# IO Parameter Builder to Nest loading and IO Assignment Generator

## Background
In brownfield projects, the project team has to take the latest controller back up as part of site survey and export the IO parameter builder and manually build the existing Nest loading and IO assignment file by hand. This activity is repetitive in nature and consumes the time of the project engineer. Since the information has lot of detail, it also opens avenues for errors.

## Solution
Build a tool that can import the IO Parameter builder (csv) and produce the nest loading and IO assignment document.

The following steps are envisaged in the generation of this tool
1. The tool parses the CSV input and produces the IO assignment - this is the most difficult part of the update as the number of tags to be updated are more requiring lot of attention to detail
2. The tool outputs the other tabs such as nest loading
3. The tool generates the reports that are present in the nest loading (like the number of spare channels)

The intent of splitting the scope into different phases will enable the progressive generation of the outputs

## Limitations
The tool is proposed to utilize only Microsoft products (Excel VBA & Power Query) as other software are not allowed to be run on work laptop.
