#---------------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      I758972
#
# Created:     17/04/2026
# Copyright:   (c) I758972 2026
# Licence:     <your licence>
#---------------------------------------------------------------------------------------


# Helper functions for Versatile TM


import os

def export_list_to_file(items, filename):
    """
    Write items (an iterable of strings) to a numbered .txt file in a predefined directory.
    Lines are numbered and the number column is aligned.
    """
    # Predefined static directory (change as needed)
    OUT_DIR = "C:/Users/I758972/Documents/automatization/clauco_test"  # example static path
    os.makedirs(OUT_DIR, exist_ok=True)

    # Ensure filename ends with .txt and is safe (basic)
    if not filename.lower().endswith(".txt"):
        filename = filename + ".txt"
    safe_name = os.path.basename(filename)  # prevents path traversal
    path = os.path.join(OUT_DIR, safe_name)

    # Convert items to list of strings
    strs = [str(x) for x in items]
    n = len(strs)
    width = len(str(n))  # digits needed for alignment

    with open(path, "w", encoding="utf-8") as f:
        for i, line in enumerate(strs, start=1):
            f.write(f"{i:>{width}}. {line}\n")

    print(f"Wrote {n} lines to: {path}")



def main():
    pass

if __name__ == '__main__':
    main()
