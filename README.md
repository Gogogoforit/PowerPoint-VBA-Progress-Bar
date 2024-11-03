# PowerPoint VBA Progress Bar

This project contains a VBA macro for Microsoft PowerPoint that automatically adds a progress bar at the bottom of each slide. The progress bar fills proportionally based on the current slide index relative to the total number of slides in the presentation.

## Features

- Creates a red progress bar on each slide.
- Progress bar width is calculated based on the slide index.
- Removes any existing progress bar before adding a new one.

## Requirements

- Microsoft PowerPoint (preferably 2016 or later).
- Basic understanding of how to use VBA in PowerPoint.

## Enabling Macros in PowerPoint

Before you can run the macro, you need to enable macros in PowerPoint. Here’s how:

1. Open PowerPoint.
2. Go to the **File** tab.
3. Click on **Options**.
4. In the **PowerPoint Options** dialog, select **Trust Center** on the left pane.
5. Click on the **Trust Center Settings** button.
6. In the **Trust Center**, select **Macro Settings**.
7. Choose **Enable all macros** (this is not recommended for security reasons; enable macros only for trusted files).
8. Click **OK** to close the Trust Center settings, then click **OK** again to close PowerPoint Options.

## Adding the VBA Macro to Your Presentation

To add the macro to your PowerPoint presentation, follow these steps:

1. Open your PowerPoint presentation.
2. Press `ALT` + `F11` to open the VBA (Visual Basic for Applications) editor.
3. In the VBA editor, go to **Insert** > **Module**. This will create a new module.
4. Copy and paste the VBA code into the module window
5. Close the VBA editor.
6. To run the macro, press ALT + F8, select AutoSections from the list, and click Run.

