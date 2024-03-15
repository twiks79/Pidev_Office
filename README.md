# Outlook Automation VBA Script

This repository contains a Visual Basic for Applications (VBA) script designed for Microsoft Outlook. The script includes several functionalities aimed at enhancing productivity and automating repetitive tasks. Key features include:

- **GetCurrentItem**: Retrieves the currently selected mail item or appointment.
- **MarkMailForMeeting**: Marks a selected mail item for a meeting.
- **AttachItem**: Attaches a mail item identified by a global EntryID to the currently selected appointment.
- **ToDo Management**: Creates tasks in specified categories ('Waiting For', 'Next', 'Small Projects', 'Large Projects', 'Later') based on the selected mail item.
- **SaveAttachment**: Saves attachments from the selected mail item to a predefined directory.
- **BulkDeleteAppointments**: Allows for bulk deletion of appointments with or without a custom cancellation message.

## Installation

To use this script, follow these steps:

1. Open Microsoft Outlook.
2. Press `Alt + F11` to open the VBA editor.
3. In the Project Explorer, right-click on `ThisOutlookSession`, and select `Insert` -> `Module`.
4. Copy and paste the provided VBA code into the new module.
5. Save the VBA project.

## Usage

To execute any of the subroutines provided by the script:

1. Ensure that the VBA script is correctly installed as per the installation instructions.
2. Go to the Developer tab in Outlook. If the Developer tab is not visible, enable it in the Outlook options.
3. Choose 'Macros', select the desired macro from the list, and click 'Run'.

Note: Some scripts operate on the currently selected mail item or appointment. Ensure you have selected the appropriate item before running these scripts.

# Powerpoint AppleScript to store text of slides in a file

## Background
You might want to use ChatBots to work on some content of your powerpoint slides. You might already have a significant amount of data in your slides but it is hard to feed slides into Chatbots.
This script will help to extract the text and support you in this task.

## Usage
1. Load this script into Apple Script Editor.
2. Open your desired powerpoint file
3. Select the slides you are interested in
4. Run the script, it will take all text on all shapes of the selected slides and save it to a specified file


# Word Macro that formats the next table in the document according to the defined standards

## Background
You might have a word document with many tables and formatting can get difficult. This macro can be used to define styling rules in the code and then always apply them to the next table.

## Usage
1. Load this macro into a word macro
2. Run the macro to change the next table below the cursor

# Remarks

## Customization

This script can be customized to fit specific workflows or requirements. Feel free to modify the code to suit your needs. Common customizations might include changing the folder paths, modifying the task categories, or adjusting the criteria for bulk deletion operations.

## Contributing

Contributions to this project are welcome! If you have suggestions for improvements or have found a bug, please open an issue or submit a pull request.

## License

This project is released under the MIT License. See the LICENSE file for details.

## Disclaimer

This script is provided "as is", without warranty of any kind. Use it at your own risk. Always back up your data before running scripts that modify your Outlook items.


