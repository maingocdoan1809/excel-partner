# excel-partner (renamed)

Excel Partner is now **Office Partner** and works in Word and PowerPoint as well as Excel. The npm package is
[`office-partner`](https://www.npmjs.com/package/office-partner).

This package only exists so that the old name keeps working. The first time you run `excel-partner` (or update an
old install), it installs `office-partner` and hands the command over to it:

```
npm install -g office-partner
office-partner setup
```

Your settings, certificate, backups and saved sign-ins are kept: Office Partner uses the same data folder
(`%LOCALAPPDATA%\Excel Partner`).
