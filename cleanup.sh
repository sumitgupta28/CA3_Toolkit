#!/usr/bin/env bash
# Cleanup script — deletes generated output files.
# Keeps: CA3_Marks/marks.xlsx (asks), students/, source scripts.

echo "CA3 Toolkit — Cleanup"
echo "----------------------------------------"

# Delete generated folders
for dir in "output" "CA3_Marks_Pdf" "__pycache__"; do
    if [ -d "$dir" ]; then
        rm -rf "$dir"
        echo "  Deleted: $dir/"
    else
        echo "  Skipped: $dir/ (not found)"
    fi
done

# Delete contents of students/ (keep the folder)
if [ -d "students" ]; then
    rm -rf students/*
    echo "  Cleared: students/"
else
    echo "  Skipped: students/ (not found)"
fi

echo "----------------------------------------"

# Ask about marks.xlsx
if [ -f "CA3_Marks/marks.xlsx" ]; then
    printf "  Delete CA3_Marks/marks.xlsx? (y/N): "
    read -r answer
    case "$answer" in
        [yY]|[yY][eE][sS])
            rm "CA3_Marks/marks.xlsx"
            echo "  Deleted: CA3_Marks/marks.xlsx"
            ;;
        *)
            echo "  Kept:    CA3_Marks/marks.xlsx"
            ;;
    esac
else
    echo "  Skipped: CA3_Marks/marks.xlsx (not found)"
fi

echo "----------------------------------------"
echo "Done."
