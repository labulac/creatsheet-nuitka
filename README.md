# creatsheet-nuitka

尝试使用nuitka打包，加快启动速度

debug : `nuitka --standalone  --mingw64 --nofollow-imports --show-memory --show-progress --plugin-enable=qt-plugins --include-qt-plugins=sensible,styles --follow-import-to=need --output-dir=outputexe creatsheet.py`

release : `nuitka --standalone --windows-disable-console --mingw64 --nofollow-imports --show-memory --show-progress --plugin-enable=qt-plugins --include-qt-plugins=sensible,styles --follow-import-to=need --output-dir=outputexe creatsheet.py`

resources文件夹需要放到dist文件夹中