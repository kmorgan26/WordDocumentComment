# Word Document Comment Consolidator

## Description
This program offers two features:
1. **Compare and Consolidate Comments**: Compares Word documents from two directories, consolidates their content and comments into a new document, and logs the differences and comments.
2. **Comments to CSV**: Extracts comments from all Word documents in a directory and writes them to a CSV file.

## Prerequisites
- .NET 8.0 Runtime
- Windows 64-bit system

## Installation
1. Download the executable file.
2. Ensure that the .NET 8.0 Runtime is installed on your system.

## Usage
1. Run the executable file.
2. Select the desired feature by entering the corresponding number:
   - `1` for Compare and Consolidate Comments
   - `2` for Comments to CSV

### Feature 1: Compare and Consolidate Comments
1. Follow the prompts to enter the paths for the following directories:
   - The first directory containing Word documents.
   - The second directory containing Word documents.
   - The output directory where the consolidated document and log files will be saved.


### Feature 2: Comments to CSV
1. Follow the prompt to enter the path to the directory containing the Word documents.
2. The program will extract comments from all Word documents in the specified directory and write them to a CSV file named `Comments.csv` in the same directory.


## Output
### Feature 1: Compare and Consolidate Comments
- A consolidated Word document containing the merged content and comments from both directories.
- A log file for each pair of compared documents, detailing the content differences and listing all comments with their authors and dates.

### Feature 2: Comments to CSV
- A CSV file named `Comments.csv` containing the comments from all Word documents in the specified directory. The CSV file includes the following columns:
  - `FileName`: The name of the Word document.
  - `Author`: The author of the comment.
  - `Date`: The date of the comment.
  - `Comment`: The text of the comment.

## Notes
- Ensure that the directories provided contain Word documents with the `.docx` extension.
- The program will create the output directory if it does not already exist.

## License
This project is licensed under the MIT License.

## Contact
For any questions or issues, please contact [Your Name] at [Your Email].
