import argparse
import csv
import sys
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Sequence, Tuple


CSV_FILE_NAME = "navisworks_views_comments.csv"
CSV_FILE_STEM = Path(CSV_FILE_NAME).stem
IMAGE_FILE_PREFIX = f"{CSV_FILE_STEM}_vp" if CSV_FILE_STEM else "vp_"


def parse_createddate(created, log):
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
        log(
            f"❌ Failed to parse createddate: {e}, raw={ET.tostring(created, encoding='unicode') if created is not None else 'None'}"
        )
        return None

def add_row(row, rows, seen, log):
    key = (row[4], row[5])  # GUID + CommentID
    if key not in seen:
        rows.append(row)
        seen.add(key)
    else:
        log(f"⚠️ Duplicate skipped: GUID={row[4]}, CommentID={row[5]}")


def recurse(folder, path, rows, seen, view_counter, log):
    folder_name = folder.attrib.get("name", "")
    new_path = path + [folder_name]

    log(f"📂 Entering folder: {' > '.join(new_path)}")

    for view in folder.findall("view"):
        view_counter[0] += 1
        view_name = view.attrib.get("name", "")
        guid = view.attrib.get("guid", "")
        image_file = f"{IMAGE_FILE_PREFIX}{str(view_counter[0]).zfill(4)}.jpg"

        log(f"  👀 Found view: {view_name} (GUID={guid}) → {image_file}")

        comments = view.find("comments")
        if comments is not None:
            for comment in comments.findall("comment"):
                cid = comment.attrib.get("id", "")
                status = comment.attrib.get("status", "")
                user = comment.findtext("user")
                body = comment.findtext("body")
                created = comment.find("createddate")
                created_str = parse_createddate(created, log) if created is not None else None

                log(f"    💬 Comment ID={cid}, Status={status}, User={user}")

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
                ], rows, seen, log)
        else:
            log(f"    ⚠️ No comments found for {view_name}")
            add_row([
                new_path[0] if len(new_path) > 0 else None,
                new_path[1] if len(new_path) > 1 else None,
                " > ".join(new_path[2:]) if len(new_path) > 2 else None,
                view_name,
                guid,
                None, None, None, None, None,
                image_file
            ], rows, seen, log)

    for child in folder.findall("viewfolder"):
        recurse(child, new_path, rows, seen, view_counter, log)


def choose_file_with_dialog() -> Optional[Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.update()
        selected = filedialog.askopenfilename(
            title="Select XML file to parse",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
        )
        root.destroy()
        if selected:
            return Path(selected)
        return None
    except Exception:
        return None


def ensure_xml_path(cli_path: Optional[str]) -> Path:
    if cli_path:
        path = Path(cli_path)
        if path.is_file():
            return path
        raise FileNotFoundError(f"XML file not found: {cli_path}")

    selected = choose_file_with_dialog()
    if selected:
        if selected.is_file():
            return selected
        raise FileNotFoundError(f"Selected file not found: {selected}")

    print("Please enter the path to the XML file:")
    user_input = input().strip().strip('"')
    if not user_input:
        raise ValueError("No XML file selected.")
    path = Path(user_input)
    if path.is_file():
        return path
    raise FileNotFoundError(f"XML file not found: {path}")


def process_xml(xml_path: Path, stream_debug: bool = False) -> Tuple[List[List[Optional[str]]], List[str]]:
    rows: List[List[Optional[str]]] = []
    debug_lines: List[str] = []
    seen: set = set()
    view_counter = [0]

    def log(message: str) -> None:
        debug_lines.append(message)
        if stream_debug:
            print(message)

    tree = ET.parse(str(xml_path))
    root = tree.getroot()

    for vf in root.findall("./viewpoints/viewfolder"):
        recurse(vf, [], rows, seen, view_counter, log)

    return rows, debug_lines


def write_outputs(rows: Sequence[Sequence[Optional[str]]], debug_lines: Sequence[str]) -> None:
    with open(CSV_FILE_NAME, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow([
            "Category",
            "Level",
            "Subfolder",
            "ViewName",
            "GUID",
            "CommentID",
            "Status",
            "User",
            "Body",
            "CreatedDate",
            "ImagePath",
        ])
        writer.writerows(rows)

    with open("debug.txt", "w", encoding="utf-8") as f:
        f.write("\n".join(debug_lines))


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Parse Navisworks XML comments into CSV")
    parser.add_argument("path", nargs="?", help="Path to the XML file to parse")
    parser.add_argument(
        "--stream-debug",
        action="store_true",
        help="Stream debug log messages to the console while processing",
    )

    args = parser.parse_args(argv)

    try:
        xml_path = ensure_xml_path(args.path)
        rows, debug_lines = process_xml(xml_path, stream_debug=args.stream_debug)
        write_outputs(rows, debug_lines)
    except Exception as exc:
        print(f"❌ Error: {exc}")
        return 1

    print("✅ Processing complete. Check navisworks_views_comments.csv and debug.txt")
    return 0


if __name__ == "__main__":
    sys.exit(main())
