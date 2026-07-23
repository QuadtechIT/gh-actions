"""Sync changed files from a GitHub release into a SharePoint document library."""

import base64
import hashlib
import os
import sys
import urllib.parse
from pathlib import Path

import msal
import requests

GRAPH = "https://graph.microsoft.com/v1.0"


def get_token() -> str:
    pfx = base64.b64decode(os.environ["SP_PUBLISHER_CERT"])
    pfx_password = os.environ.get("SP_PUBLISHER_CERT_PASSWORD", "").encode() or None
    from cryptography.hazmat.primitives.serialization import pkcs12, Encoding, PrivateFormat, NoEncryption

    key, cert, _ = pkcs12.load_key_and_certificates(pfx, pfx_password)
    private_pem = key.private_bytes(
        Encoding.PEM, PrivateFormat.PKCS8, NoEncryption()
    ).decode()
    thumbprint = cert.fingerprint(__import__("cryptography.hazmat.primitives.hashes", fromlist=["SHA1"]).SHA1()).hex()

    app = msal.ConfidentialClientApplication(
        client_id=os.environ["AZURE_CLIENT_ID"],
        authority=f"https://login.microsoftonline.com/{os.environ['AZURE_TENANT_ID']}",
        client_credential={"private_key": private_pem, "thumbprint": thumbprint, "public_certificate": cert.public_bytes(Encoding.PEM).decode()},
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise SystemExit(f"Token acquisition failed: {result.get('error_description')}")
    return result["access_token"]


class SharePoint:
    def __init__(self, token: str):
        self.s = requests.Session()
        self.s.headers["Authorization"] = f"Bearer {token}"

    def _check(self, r, context):
        if not r.ok:
            raise SystemExit(f"{context} failed [{r.status_code}]: {r.text[:400]}")
        return r

    def resolve_drive(self, hostname: str, site_path: str, library: str) -> str:
        r = self._check(
            self.s.get(f"{GRAPH}/sites/{hostname}:{site_path}"), "Site lookup"
        )
        site_id = r.json()["id"]
        r = self._check(self.s.get(f"{GRAPH}/sites/{site_id}/drives"), "Drive list")
        for d in r.json()["value"]:
            if d["name"].lower() == library.lower():
                return d["id"]
        names = ", ".join(d["name"] for d in r.json()["value"])
        raise SystemExit(f"Library '{library}' not found. Available: {names}")

    def remote_hash(self, drive_id: str, path: str):
        """Return SharePoint's quickXorHash for a file, or None if absent."""
        enc = urllib.parse.quote(path)
        r = self.s.get(f"{GRAPH}/drives/{drive_id}/root:/{enc}")
        if r.status_code == 404:
            return None
        self._check(r, "Item lookup")
        return r.json().get("file", {}).get("hashes", {}).get("quickXorHash")

    def upload(self, drive_id: str, path: str, local: Path):
        enc = urllib.parse.quote(path)
        data = local.read_bytes()
        if len(data) < 4 * 1024 * 1024:
            r = self.s.put(
                f"{GRAPH}/drives/{drive_id}/root:/{enc}:/content",
                data=data,
                headers={"Content-Type": "application/octet-stream"},
            )
            self._check(r, f"Upload {path}")
            return r.json()["id"]
        # Large file: upload session
        r = self._check(
            self.s.post(f"{GRAPH}/drives/{drive_id}/root:/{enc}:/createUploadSession",
                        json={"item": {"@microsoft.graph.conflictBehavior": "replace"}}),
            f"Upload session {path}",
        )
        url = r.json()["uploadUrl"]
        chunk = 5 * 1024 * 1024
        for start in range(0, len(data), chunk):
            end = min(start + chunk, len(data)) - 1
            cr = requests.put(
                url,
                data=data[start:end + 1],
                headers={"Content-Range": f"bytes {start}-{end}/{len(data)}"},
            )
            if cr.status_code not in (200, 201, 202):
                raise SystemExit(f"Chunk upload failed [{cr.status_code}]: {cr.text[:400]}")
        return cr.json()["id"]

    def stamp(self, drive_id: str, item_id: str, tag: str, sha: str):
        """Write release metadata as list item columns. Non-fatal if columns absent."""
        r = self.s.patch(
            f"{GRAPH}/drives/{drive_id}/items/{item_id}/listItem/fields",
            json={"ReleaseTag": tag, "CommitSHA": sha[:12]},
        )
        if not r.ok:
            print(f"  ! metadata stamp skipped [{r.status_code}] — check columns exist")

    def delete(self, drive_id: str, path: str):
        enc = urllib.parse.quote(path)
        r = self.s.delete(f"{GRAPH}/drives/{drive_id}/root:/{enc}")
        if r.status_code == 404:
            print(f"  · already absent: {path}")
        elif not r.ok:
            raise SystemExit(f"Delete {path} failed [{r.status_code}]: {r.text[:400]}")


def quick_xor(data: bytes) -> str:
    """SharePoint's quickXorHash. Used to skip no-op uploads."""
    width, shift, buf = 160, 11, bytearray(20)
    bit = 0
    for i in range(0, len(data), 1):
        idx = (bit // 8) % 20
        buf[idx] ^= data[i]
        bit = (bit + shift) % width
    length = len(data).to_bytes(8, "little")
    for i in range(8):
        buf[(width // 8) - 8 + i] ^= length[i]
    return base64.b64encode(bytes(buf)).decode()


def main(changes_file: str):
    src_root = os.environ["SOURCE_DIR"].strip("/")
    tgt_root = os.environ["TARGET_ROOT"].strip("/")
    tag = os.environ["RELEASE_TAG"]
    sha = os.environ["COMMIT_SHA"]

    sp = SharePoint(get_token())
    drive_id = sp.resolve_drive(
        os.environ["SITE_HOSTNAME"], os.environ["SITE_PATH"], os.environ["LIBRARY"]
    )
    print(f"Drive resolved. Publishing {tag} under /{tgt_root}/")

    def remote_path(repo_path: str) -> str:
        if src_root and repo_path.startswith(src_root + "/"):
            rel = repo_path[len(src_root) + 1:]
        else:
            rel = repo_path
        return f"{tgt_root}/{rel}"

    uploaded = skipped = deleted = 0

    with open(changes_file) as fh:
        for line in fh:
            line = line.rstrip("\n")
            if not line:
                continue
            parts = line.split("\t")
            status = parts[0][0]

            if status == "R":                     # rename: delete old, upload new
                old, new = parts[1], parts[2]
                sp.delete(drive_id, remote_path(old))
                print(f"  - {remote_path(old)} (renamed)")
                deleted += 1
                status, path = "M", new
            else:
                path = parts[1]

            target = remote_path(path)

            if status == "D":
                sp.delete(drive_id, target)
                print(f"  - {target}")
                deleted += 1
                continue

            local = Path(path)
            if not local.is_file():
                print(f"  ! missing locally, skipping: {path}")
                continue

            if sp.remote_hash(drive_id, target) == quick_xor(local.read_bytes()):
                print(f"  = {target} (identical, skipped)")
                skipped += 1
                continue

            item_id = sp.upload(drive_id, target, local)
            sp.stamp(drive_id, item_id, tag, sha)
            print(f"  + {target}")
            uploaded += 1

    print(f"\nDone. {uploaded} uploaded, {skipped} unchanged, {deleted} removed.")


if __name__ == "__main__":
    main(sys.argv[1])
