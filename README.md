# Powershell Helper Scripts
For tasks that are done often or rarely it can be useful to wrap up some quirky or verbose commands into shorter custom functions that are easier to remember.

# Install

1. Download the folder and place it in your own module folder. `$env:PSModulePath.split(';')` will show you these locations.
1. Ensure that folder is named `Helper`

Because the module is in your known path, powershell can find the commands by name to automatically import as you type them.

## Manual use
1. When needed, import the module with `Import-module <PathToFolder>\helper.psd1`

# How To Use

* `get-command -module helper` will show you the commands from this module
* `get-help <command name>` will print the help of a command you are interested in.
* `get-help <command name> -examples` will display how you can use a command
* `get-help <command name> -full` you will find everything there is to know.
* `get-help about_helper` will show you the module general help file
