# copy_Excel
## 要求仕様
.xlsxファイルをロードし、一回のmain.pyの実行で一行クリップボードにコピー

何行めまでコピーしたか保存しておき、次回の実行では次の行をクリップボードにコピー

## How to run(Power Shell)<初回のみ>
```PowerShell
git clone https://github.com/shiba0627/copy_Excel.git
cd copy_Excel
python -m venv .venv_ex
.\.venv_ex\Scripts\activate 
python -m pip install -r requirements.txt 
python main.py
```
## How to run(Power Shell)<二回目以降>
```PowerShell
cd copy_Excel
.\.venv_ex\Scripts\activate 
python main.py
```
