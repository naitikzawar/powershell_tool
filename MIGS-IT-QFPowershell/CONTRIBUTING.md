# Contribute
---
Contributions of code are welcome!
Please create a new branch to upload your code then create a pull request. Please do NOT push directly into the main branch!

[Please use the same process as for the MIGS-IT-CustomerSolutions repo.](https://confluence.derivco.co.za/display/MIGS/Contributing+to+the+MIGS-IT-CustomerSolutions+Repository)

The process is outlined below:
1. Navigate to the repository in **Azure DevOps** and select the *Branch* dropdown. This will probably be showing the `main` branch when you first open the page.
2. Select *+ New Branch* under the Branch dropdown menu.
3. Enter a branch name (e.g: Test, Demo, MyNewFeature etc) and set *Based on* to `main`
4. Once your branch is created, you will see your new branch name displayed on the branch dropdown. 
5. If you used Git to clone the main repository, you can switch to your new branch as follows:
* `cd` into the repository folder.
* Run `git pull`
* run `git checkout` followed by the name of the new branch. e.g. `git checkout MyNewFeature`
* You will now be working locally in the new branch.
5. Make the required changes on your branch, and commit the changes by selecting the commit button in the Azure DevOps page, or using Git to commit and push to the branch:
* `cd` into the repository folder.
* Run `git add -A` to stage all changed, added or deleted files.
* Run `git commit` to commit the changes locally. You will probably be asked to enter a short comment about your changes.
* Run `git push` to upload your committed changes into the Azure Devops repository, under the new branch.
6. After you have committed your changes, select the pull request icon from the left-hand side panel in Azure DevOps.
7. Select *New Pull Request*.
8. When on the pull request page, ensure that you are merging your branch into the main branch, enter a title and description, and select the *create* button. 
9. [You can optionally set the pull request to autocomplete](https://learn.microsoft.com/en-us/azure/devops/boards/work-items/auto-complete-work-items-pull-requests?view=azure-devops)
10. Your pull request will now be active and pending approval. Please notify the maintainers to review and approce the request (Refer to the *Contact and Help* section in the `README.md` file)
11. Once your pull request has been approved, select the *complete* drop-down and click *complete* again. 
12. Ensure that the *Delete Branch after merging* option is selected before completing the pull request.


We recommend using ``VSCode``, with the Powershell extension. VSCode includes Git source control out of the box, and makes Git operations like commit, push, pull etc from the Azure DevOps repo quite simple.


# Modules
---
This repository contains several PowerShell module files.

The primary module definition file is ``QuickFire.psd1`` in the root folder of the repository. This is the file that you should import into PowerShell to access all the functions and cmdlets.
When it is imported, this module definition file automatically loads all the ``*.psm1`` files located in  the ``src`` folder.

You may wish to add your own module file to the project rather than edit one of the existing module files.
The process for this is as follows:

1. Create your new ``*.psm1`` module file under the ``src`` folder.
2. Edit the ``QuickFire.psd1`` in the root folder of the repository.
3. Locate the ``NestedModules`` array.
4. Add your new file name (including the parent ``src`` folder) to the ``NestedModules`` array. Note the comma on the end of each line except the last.
5. Save the file. Your new module file should be loaded when ``QuickFire.psd1`` is imported into PowerShell. You can confirm with: `Get-Command -module Quickfire`


For example:
``` 
NestedModules = @(
    'src\Quickfire-SD.psm1',
    'src\Quickfire-SQL.psm1',
    'src\Quickfire-Gamestats.psm1',
    'src\Quickfire-Playcheck.psm1',
    'src\Quickfire-MyNewModule.psm1'
    )

``` 

Note that ``src\Quickfire.psm1`` is the `root` module and is loaded by the ``RootModule`` line; it does not need to be added to the ``NestedModules`` array.


# Style and Format rules
---
While we are not particularly strict on styling and formatting, please try to follow the [Unofficial PowerShell Best Pracites and Style Guide](https://github.com/PoshCode/PowerShellPracticeAndStyle) wherever possible. An `.editorconfig` file is included in the repo to configure some of the below settings.

General guidelines:
* Functions should have a documentation section included, e.g. Synoposis, Description, Parameters, Examples.
* Use PascalCase for function and cmdlet names, variables and constants etc. Keywords (eg foreach) and operators (eg -or, -and, -eq) should be in lowercase.
* Indent your functions and statements. Indents should be set to 4 spaces.
* Line endings should be Unix format (LF) and not Windows format (CRLF).
* Curly braces (e.g. for ScriptBlocks) should open on the end of the previous line. The closing brace should be on a seperate line by itself; unless you are only running one command, in which case the whole ScriptBlock can be on a single line including its braces.


# Work Tracking
---
The MIGS Customer Solutions team uses an [Azure DevOps board](https://dev.azure.com/Derivco/Software/_boards/board/t/MIGS%20-%20IT%20-%20Customer%20Solutions%20Service%20Team/Stories) to track our Reboot activities. You may lodge a task or work item here.


# Permissions
---
Some basic functions of the script will work without any special permissions.

Access to the Azure DevOps repository, PlayCheck and Game Statistics sites, and the QuickFire SQL databases are restricted to members of the MIGS team. Please contact a Derivco MIGS Service Owner if you require access, which will be granted at the Service Owner's discretion.


# Issues and Bug Reports
---
Please refer to the *Contact and Help* section of README.md