## Usage

1. Download the latest [release](https://github.com/therepos/excelUI/releases/excelUI/latest/download/excelUI.zip). 
2. Install with excelUIsetup.cmd
3. The addin will be located at `%APPDATA%\Microsoft\AddIns`.

## Updating the add-in

Edit the `.bas` / `.cls` / `.xml` files in `src/excelUI` and push to `main`. CI rebuilds `excelUI.xlam` from source
(`scripts/vbabuild.py`, no Excel needed), commits it back and publishes a release. Don't edit the `.xlam` directly.

## Structure

```
excelUI
├── .github
│   └── workflows
├── .scraps
├── app                 # Executables
├── docs                # Documentation
├── scripts             
│   └── docgen-bas.py   # Autogenerate doc/bas.md
├── src
│   └── bas             # Documentation
└── tmp                 # Temporary files
```

## Resources

- [Office RibbonX Editor](https://github.com/fernandreu/office-ribbonx-editor)
- [ImageMSO Gallery](https://bert-toolkit.com/imagemso-list.html)

