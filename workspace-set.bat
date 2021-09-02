mkdir .vscode
cd .vscode
echo { > settings.json
echo     "code-runner.showRunIconInEditorTitleMenu": true, >> settings.json
echo     "code-runner.runInTerminal": true, >> settings.json
echo     "terminal.integrated.fontSize": 16, >> settings.json
echo     "files.encoding": "shiftjis" >> settings.json
echo } >> settings.json
