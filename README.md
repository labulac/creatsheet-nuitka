# creatsheet-nuitka

����ʹ��nuitka������ӿ������ٶ�

debug : `nuitka --standalone  --mingw64 --nofollow-imports --show-memory --show-progress --plugin-enable=qt-plugins --include-qt-plugins=sensible,styles --follow-import-to=need --output-dir=outputexe creatsheet.py`

release : `nuitka --standalone --windows-disable-console --mingw64 --nofollow-imports --show-memory --show-progress --plugin-enable=qt-plugins --include-qt-plugins=sensible,styles --follow-import-to=need --output-dir=outputexe creatsheet.py`

resources�ļ�����Ҫ�ŵ�dist�ļ�����