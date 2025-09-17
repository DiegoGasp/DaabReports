import xml.etree.ElementTree as ET
import csv
from datetime import datetime

tree = ET.parse(r"C:\Users\Diego\daabcontech.com\PROJECTS - Documents\250909 - 1717 N Flagler\02_WIP\02.08 ISSUE REPORTS\XML\1717 N Flagler - Coordination Model_2025 - Abner.xml")
root = tree.getroot()

rows = []
debug_lines = []
seen = set()  # Track unique (GUID, CommentID) pairs
view_counter = 0  # Global counter for vp####.jpg

def parse_createddate(created):
    """Convert <createddate><date .../> into yyyy/mm/dd"""
    try:
        date_node = created.find("date") if created is not None else None
        if date_node is None:
            return None

        year = int(date_node.attrib.get("year", 0))
        month = int(date_node.attrib.get("month", 0))
        day = int(date_node.attrib.get("day", 0))

        if year < 1900 or month == 0 or day == 0:
            return None

        dt = datetime(year, month, day)
        return dt.strftime("%Y/%m/%d")

    except Exception as e:
        debug_lines.append(f"❌ Failed to parse createddate: {e}, raw={ET.tostring(created, encoding='unicode') if created is not None else 'None'}")
        return None

def add_row(row):
    key = (row[4], row[5])  # GUID + CommentID
    if key not in seen:
        rows.append(row)
        seen.add(key)
    else:
        debug_lines.append(f"⚠️ Duplicate skipped: GUID={row[4]}, CommentID={row[5]}")

def recurse(folder, path):
    global view_counter

    folder_name = folder.attrib.get("name", "")
    new_path = path + [folder_name]

    debug_lines.append(f"📂 Entering folder: {' > '.join(new_path)}")

    for view in folder.findall("view"):
        view_counter += 1
        view_name = view.attrib.get("name", "")
        guid = view.attrib.get("guid", "")
        image_file = f"vp{str(view_counter).zfill(4)}.jpg"

        debug_lines.append(f"  👀 Found view: {view_name} (GUID={guid}) → {image_file}")

        comments = view.find("comments")
        if comments is not None:
            for comment in comments.findall("comment"):
                cid = comment.attrib.get("id", "")
                status = comment.attrib.get("status", "")
                user = comment.findtext("user")
                body = comment.findtext("body")
                created = comment.find("createddate")
                created_str = parse_createddate(created) if created is not None else None

                debug_lines.append(f"    💬 Comment ID={cid}, Status={status}, User={user}")

                add_row([
                    new_path[0] if len(new_path) > 0 else None,
                    new_path[1] if len(new_path) > 1 else None,
                    " > ".join(new_path[2:]) if len(new_path) > 2 else None,
                    view_name,
                    guid,
                    cid,
                    status,
                    user,
                    body,
                    created_str,
                    image_file
                ])
        else:
            debug_lines.append(f"    ⚠️ No comments found for {view_name}")
            add_row([
                new_path[0] if len(new_path) > 0 else None,
                new_path[1] if len(new_path) > 1 else None,
                " > ".join(new_path[2:]) if len(new_path) > 2 else None,
                view_name,
                guid,
                None, None, None, None, None,
                image_file
            ])

    for child in folder.findall("viewfolder"):
        recurse(child, new_path)

# ✅ Start only at top-level categories
for vf in root.findall("./viewpoints/viewfolder"):
    recurse(vf, [])

# Save main data to CSV
with open("navisworks_views_comments.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow([
        "Category", "Level", "Subfolder",
        "ViewName", "GUID",
        "CommentID", "Status", "User", "Body", "CreatedDate",
        "ImagePath"
    ])
    writer.writerows(rows)

# Save debug log
with open("debug.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(debug_lines))

print("✅ Processing complete. Check navisworks_views_comments.csv and debug.txt")
