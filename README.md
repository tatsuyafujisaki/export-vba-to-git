# Prerequisites
1. Open Excel
2. Menu bar > `File` > `Options` > `Trust Center` > `Trust Center Settings` > `Macro Settings` > `Trust access to the VBA project object model` > Check

# Preparation
1. Let's say you want to export VBA of `foo.xlsm` to `https://git.example.com/foo.git`
2. Download `template.cmd` and `template.vbs` to the directory containing `foo.xlsm`
3. Rename the downloaded files to `foo.cmd` and `foo.vbs`
4. Set `https://example.com/foo.git` to the variable `gitRepositoryUrl` in `foo.cmd`

# How to use
1. Double-click `foo.cmd`
