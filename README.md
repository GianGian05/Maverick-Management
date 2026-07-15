# Maverick Roof RFP & Bid-Leveling MVP

This prototype automates the first roof procurement workflow:

1. Enter project information once in `templates/Roof_Project_Intake.xlsx`.
2. Generate a project-specific Word RFP and standardized contractor Excel bid form.
3. Send the Excel bid form to contractors.
4. Select one to four returned Excel files.
5. Generate a populated bid-leveling workbook with exclusions and incomplete responses flagged.

## First-time setup

Windows users should double-click `Setup_Windows.bat`. Mac users should double-click `Setup_Mac.command`. Python 3 and Microsoft Word should be installed. Microsoft Word is used only for automatic PDF conversion; the DOCX output works without it.

## Normal use

Windows: double-click `Start_Maverick_Roof_Tool.bat`.

Mac: double-click `Start_Maverick_Roof_Tool.command`.

Choose the intake workbook, then use the two buttons in order:

- **Generate RFP + Contractor Bid Form**
- **Select Excel Bids + Create Leveling**

## Controls and safeguards

- The approved 47-page RFP is stored as a locked master template.
- The tool checks the template checksum and stops if the master was altered.
- Only approved project placeholders are changed.
- Project state/license mismatches, missing required fields, impossible dates, and invalid roof area stop generation.
- Contractor sheet names and rows must not be renamed or deleted.
- Missing bid values, blank scope responses, and exclusions are surfaced in the leveling workbook.
- The `Import Audit` sheet records the file, bidder, template version, imported-field count, warnings, and readiness status.

## Current version boundary

Version 1.0 supports the Maverick roof-recover structure represented by the Minneapolis sample: two base-bid options, one warranty alternate, six unit-price alternates, and 46 standardized scope-compliance items. Paving, structural, ground-up construction, assessment-report photo automation, and OM-to-pro-forma extraction are intentionally outside this release.

## Updating the legal template

Do not edit `templates/Roof_RFP_Master_Template.docx`. When Maverick approves a legal or administrative change, create a new template version and regenerate `templates/RFP_Template_Manifest.json`. This keeps project users from accidentally changing locked language.
