import re
import os

# Modify this to point to your markdown file
markdown_file_path = "input/PASTE_YOUR_PROJECT_HERE.md"
project_root = "project-root"

# Regex to match the pattern:
# 1. A line starting with ### followed by space(s), then a backtick, the filename, another backtick
# 2. Followed by a code block starting with ```<lang> and ending with ```
pattern = re.compile(
    r'^###\s+`([^`]+)`.*?\n```([a-zA-Z0-9]+)\n(.*?)```',
    re.DOTALL | re.MULTILINE
)

# Create project root directory if it doesn't exist
if not os.path.exists(project_root):
    os.makedirs(project_root, exist_ok=True)

with open(markdown_file_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Find all matches
matches = pattern.findall(content)

# matches will be a list of tuples: (filename, language, code)
# Example: ("package.json", "json", "{\n  \"name\": ... }")

for filename, language, code in matches:
    # Construct the full path for the output file
    full_path = os.path.join(project_root, filename)
    
    # Ensure directories exist if filename includes paths like "src/taskpane/taskpane.html"
    dir_name = os.path.dirname(full_path)
    if dir_name and not os.path.exists(dir_name):
        os.makedirs(dir_name, exist_ok=True)

    # Write the code snippet to the file
    with open(full_path, 'w', encoding='utf-8') as outfile:
        outfile.write(code.strip() + '\n')

    print(f"Extracted: {full_path} ({language})")
