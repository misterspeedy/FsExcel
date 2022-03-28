# Contributing

The FsExcel repo has a few unusual features, so please give this a read before diving in to make changes.

## Code of Conduct

Please aim to leave anyone you interact with happier than they were before you interacted with them.

## Technical Setup

You may find it easiest to work with this codebase using VS Code. The instructions below assume this.

The tutorial notebook `Tutorial.dib` has three important roles:

1) It gives users a downloadable resource which they can run to explore all FsExcel features.

1) It is used to generate the `README.md` file which allows users to view these features in `GitHub`, together with expected results.

1) It is used to generate a suite of regression tests.

To make all this work requires a little setup.

- You will need to enable run-on-folder-open. To do this:

    - Open the command palette (SHIFT+CTRL+P) and choose "Tasks: Manage Automatic Tasks in Folder"

    - Choose "Allow Automatic Tasks in Folder".  Note that this setting is *not* in the usual VS Code settings JSON files.

    - Close and re-open the workspace.
    
    - At this point you should see three items named "Start watching..." in your terminal list.  This is the three dotnet 'watches' which automatically update the `README.md` and the regression tests, based on `Tutorial.dib`.

- To double check that all is working, make a trivial change to `Tutorial.dib` and save. After a moment you should see *two* changes appear in the source control panel, one in `Tutorial.dib` and one in `README.md`. Reverse and save your trivial change and both these file changes should disappear.

*Don't edit `README.md` directly. Always change it by editing `Tutorial.dib` and saving.*

It's not necessary for most development work, but to explore further how this aspect of the repo works, look at `tasks.json` in the repo root, and at the `.fsx` scripts which it calls.

## Regression Tests

As detailed above, the regression tests are generated based on the spreadsheets generated in `Tutorial.dib`. The process overall is:

1) You add or change some feature in FsExcel.

1) You make appropriate changes in `Tutorial.dib` to demonstrate the feature.

1) When you save `Tutorial.dib`, a dotnet watch uses `DibToActualsScript.fsx` to generate a new script called `CreateRegressionTestActuals.fsx`. The content of `CreateRegressionTestActuals.fsx` is essentially the source from each of the F# cells in the tutorial.

1) Another dotnet watch runs this newly-generated script to generate spreadsheets in `src/Tests/Actual`.

1) When you run the tests using `dotnet test`, the regression test compares every spreadsheet in `src/Tests/Expected` with those in `src/Tests/Actual` and reports an error where these differ.

Note that the regression tests don't yet compare every cell attribute that can be set using FsExcel. This is outstanding work.

## Developer Workflow

With this in mind you'll need to take the following steps when making changes:

1) If you are *adding* a feature, you'll need to demonstrate it in one or more F# cells (with appropriate markdown commentary) in `Tutorial.dib`. The last line of markdown before the
new F# cell should look like this:
```html
<!-- Test -->
```
This will not be visible in the rendered markdown, but tells the regression test generator script to include the code from the following F# cell in the test actuals generator.

2) If you are *amending* a feature and the change affects expected output, you'll need to amend existing F# cells and commentary in `Tutorial.dib`.

3) In order to run against the locally-built version of FsExcel you'll need to *temporarily* add the following lines to the beginning of each affected notebook cell:

```fsharp
#r "nuget: ClosedXML"
#r "../FsExcel/bin/Debug/netstandard2.1/FsExcel.dll"
let savePath = "/temp"
```

4) You'll need to reset the notebook kernel each time you want to `dotnet build` the local version (CTRL+SHIFT+P -> `.NET Interactive: restart the current notebook's kernel`). This is because running the cell locks the object code.

1) Once your change is successfully generating a spreadsheet, take a carefully cropped screenshot and include it as an example after the code cell. (See existing markdown for many examples.) Your screenshot won't show in the markdown until you have pushed your changes, as the links are into GitHub.

1) When you are happy with the change, you'll need to copy any new or changed output spreadsheets from the `savePath` where the notebook would have written them, into `src/Tests/Expected`.

1) Now run the tests with `dotnet test`.

With the feature working and regression tests passing...

1) Remove the three lines you added to `Tutorial.dib` cells to run the local copy.

1) Verify in your git changes that the edits you made in `Tutorial.dib` are reflected in `README.md`.

1) Also in your git changes, check that any new/edited spreadsheets in `src/Tests/Actual`, and any screenshots, are reflected.

1) If appropriate, upversion in `FsExcel.fsproj` to ensure a new Nuget package is built.

1) Push your changes!

1) Once the package has been published on `nuget.org`, run the entire notebook again to ensure the changes work correctly from the published version.


