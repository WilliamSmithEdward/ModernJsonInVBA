"""Pin the latest verified YARA Forge core release for the security scan."""

import hashlib
from io import BytesIO
import json
import os
from pathlib import Path
import re
import urllib.request
import zipfile


PIN = Path(__file__).with_name("yara-forge.json")
API_URL = "https://api.github.com/repos/YARAHQ/yara-forge/releases/latest"
ASSET_NAME = "yara-forge-rules-core.zip"


def fetch(url):
    headers = {"Accept": "application/vnd.github+json", "User-Agent": "ModernJsonInVBA-security"}
    token = os.environ.get("GITHUB_TOKEN")
    if token and url.startswith("https://api.github.com/"):
        headers["Authorization"] = f"Bearer {token}"
    with urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=60) as response:
        return response.read()


def candidate(release):
    tag = release["tag_name"]
    if not re.fullmatch(r"\d{8}", tag) or release["draft"] or release["prerelease"]:
        raise ValueError("latest YARA Forge release is not a stable dated release")
    assets = [asset for asset in release["assets"] if asset["name"] == ASSET_NAME]
    if len(assets) != 1:
        raise ValueError("expected exactly one YARA Forge core archive")
    asset = assets[0]
    url = f"https://github.com/YARAHQ/yara-forge/releases/download/{tag}/{ASSET_NAME}"
    if asset["browser_download_url"] != url:
        raise ValueError("YARA Forge core archive URL does not match its release")
    digest = asset.get("digest", "")
    if not re.fullmatch(r"sha256:[0-9a-f]{64}", digest):
        raise ValueError("YARA Forge core archive has no valid SHA-256 digest")
    return tag, url, digest.removeprefix("sha256:")


def main():
    current = json.loads(PIN.read_text(encoding="utf-8"))
    release = json.loads(fetch(API_URL))
    tag, url, expected_hash = candidate(release)
    if tag == current["tag"]:
        if expected_hash != current["sha256"]:
            raise ValueError("published YARA Forge checksum changed for the pinned release")
        print(f"YARA Forge {tag} is already pinned")
        return
    if tag < current["tag"]:
        raise ValueError("latest YARA Forge release predates the pinned release")
    archive = fetch(url)
    actual_hash = hashlib.sha256(archive).hexdigest()
    if actual_hash != expected_hash:
        raise ValueError("downloaded YARA Forge archive does not match its published SHA-256")
    with zipfile.ZipFile(BytesIO(archive)) as zipped:
        names = [name for name in zipped.namelist() if name.endswith("/yara-rules-core.yar")]
        if len(names) != 1:
            raise ValueError("YARA Forge core archive does not contain exactly one core rule file")
    PIN.write_text(json.dumps({"tag": tag, "sha256": actual_hash}, indent=2) + "\n", encoding="utf-8")
    print(f"Pinned YARA Forge {tag} at SHA-256 {actual_hash}")


if __name__ == "__main__":
    main()
