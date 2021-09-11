
# vbacGUI

a GUI Frontend of the package including vbac & pt for VBA-er

## Required Enviroment

- Windows 7, 8
- .NET Framework 4.5

## Usage
### Decompile (ソースコードのエクスポート)

1. binaryフォルダ（デフォルトでは、binフォルダ）に
　　VBAコード（マクロ）を含むWord, Excel, Accessのバイナリファイルを
　　入れてください。
2. 「Decombine」をクリック。
3. sourceフォルダ(デフォルトでは、srcフォルダ)に、ソースコードがエクスポートされます。

### Compile (ソースコードのインポート)
1. sourceフォルダに、ソースコードが入れます。
2. 「Combine」をクリック。
3. binフォルダに、ソースコードがインポートされたバイナリファイルが作成されます。

### Option
詳しくは、vbacのOptionを参照してください。

# acknowledgement
- vbac.wsf is created by @igeta
-- https://github.com/vbaidiot/Ariawase
- pt.exe is created by @monochromegane
-- https://github.com/monochromegane/the_platinum_searcher
