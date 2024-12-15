
# Markdown to Codebase Generator

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Python Version](https://img.shields.io/badge/python-3.6%2B-blue)
![Build Status](https://img.shields.io/badge/build-passing-brightgreen)

A powerful Python utility that automatically converts AI chatbot code responses into a fully structured codebase.

## 🚀 Overview

 Markdown to Codebase Generator streamlines the process of converting code snippets from AI chatbots (like o1 Pro, o1, GPT 4o, Claude 3.5 Sonnet or Gemini 2.0 Flash) into a complete, organized project structure. Simply paste the chatbot's response into the markdown file at `input/PASTE_YOUR_PROJECT_HERE.md`, run `main.py` and the tool will automatically generate your entire structured codebase in the `/project-root` directory with the proper structure and file types.

## ✨ Features

- **One-Click Codebase Generation**: Convert markdown-formatted code into a complete project structure
- **AI Chatbot Compatible**: Works with responses from OpenAI's ChatGPT, Anthropic's Claude, Google's Gemini, and more
- **Automatic Directory Structure**: Creates nested directories and files based on the markdown structure
- **Language Agnostic**: Supports any programming language or file type
- **Preserves Code Formatting**: Maintains original code indentation and structure

## 🛠️ Installation

```bash
# Clone the repository
git clone https://github.com/Slyycooper/Markdown-to-Codebase-Generator.git

# Navigate to the project directory
cd Markdown-to-Codebase-Generator

# Create input/output directories
mkdir -p input output
```

## 📖 Usage

1. Create a markdown file named `PASTE_YOUR_PROJECT_HERE.md` in the `input` directory
2. Paste your the chatbot's response using this format (typically the chatbot's response will be in this format already):

```markdown
### `filename.ext`
```language
// Your code here
```
```markdown
### `tools/filename.ext`
```language
// Your code here
```
```markdown
### `src/taskpane/filename.ext`
```language
// Your code here
```




3. Run the script:

```bash
python main.py
```

4. Find your generated codebase in `output/project-root`

## 📁 Directory Structure

```
codebase-generator/
├── input/
│   └── PASTE_YOUR_PROJECT_HERE.md
├── output/
│   └── project-root/
├── main.py
└── README.md
```

## 🌟 Example

Input (`PASTE_YOUR_PROJECT_HERE.md`):
```markdown
---

### `src/taskpane/scanExcelApi.js`

```javascript
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var fs = require("fs");

function scanExcelApi() {
  var excelApiMethods = [];
  for (var key in Excel) {
    if (Excel.hasOwnProperty(key)) {
      var property = Excel[key];
      if (typeof property === "function") {
        excelApiMethods.push("Function: " + key);
      } else if (typeof property === "object") {
        excelApiMethods.push("Object: " + key);
        for (var subKey in property) {
          if (property.hasOwnProperty(subKey) && typeof property[subKey] === "function") {
            excelApiMethods.push("  - Method: " + subKey);
          }
        }
      }
    }
  }
  fs.writeFileSync("excelApiMethods.txt", excelApiMethods.join("\n"), "utf8");
  console.log("Excel API methods have been written to excelApiMethods.txt");
}
scanExcelApi();
```

---



Output:
```
project-root/
├── src/
    └── taskpane/
       └── scanExcelApi.js
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Inspired by the need to streamline the world around us
- Thanks to all platforms making code generation accessible

## 📬 Contact

Trevor Nichols - [@trevnics](https://x.com/trevnics)

Project Link: [https://github.com/Slyycooper/codebase-generator](https://github.com/Slyycooper/codebase-generator)
