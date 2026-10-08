import re
import subprocess

version = input("Version (example 14.05): ").strip()
message = input("Notice (example: Coming soon): ").strip()

if not version or not message:
    print("Version and notice are required.")
    raise SystemExit(1)

print("\nEnter the Update Package items.")
print("Press Enter on a blank line when finished.\n")

items = []
while True:
    item = input(f"Item {len(items) + 1}: ").strip()
    if not item:
        break
    items.append(item)

if not items:
    print("At least one update package item is required.")
    raise SystemExit(1)

path = "index.html"

with open(path, "r", encoding="utf-8") as f:
    html = f.read()

# Update version / notice
version_pattern = re.compile(
    r'(<p style="text-align:center; margin: 0 0 8px 0;[^>]*>\s*)'
    r'Phillip Accounting v.*?'
    r'(\s*</p>)',
    re.DOTALL
)

version_replacement = (
    r'\1Phillip Accounting v' + version +
    r' — ' + message +
    r'\2'
)

html, count = version_pattern.subn(version_replacement, html, count=1)

if count == 0:
    print("Could not find the version line.")
    raise SystemExit(1)

# Update package bullets
items_html = "\n".join(
    f'    <li>{item}</li>' for item in items
)

package_pattern = re.compile(
    r'(<b>Update Package:</b>\s*'
    r'<ul[^>]*>\s*)'
    r'.*?'
    r'(\s*</ul>)',
    re.DOTALL
)

package_replacement = (
    r'\1' + items_html + r'\2'
)

html, count = package_pattern.subn(package_replacement, html, count=1)

if count == 0:
    print("Could not find the Update Package section.")
    raise SystemExit(1)

with open(path, "w", encoding="utf-8") as f:
    f.write(html)

subprocess.run(["git", "add", "index.html", "update_notice.py"], check=True)
subprocess.run(
    ["git", "commit", "-m", f"Update website package to v{version}"],
    check=True
)
subprocess.run(["git", "push", "origin", "main"], check=True)

print(f"\nWebsite updated: Phillip Accounting v{version} — {message}")
print(f"Update Package: {len(items)} items")
print("Render will deploy automatically.")
