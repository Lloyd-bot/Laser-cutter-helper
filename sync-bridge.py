# Copy print-to-laser.ps1 into index.html, so the tool can hand out the script
# itself and stay the one file you carry to the laser PC.
#
# print-to-laser.ps1 is the original. The copy inside index.html is just that -
# a copy - and it goes stale the moment the script is edited without running
# this. So: edit the script, then run
#
#     python sync-bridge.py
#
# and commit both files together. Run it with --check to be told whether they
# have drifted apart without changing anything (exit code 1 if they have).
import io, os, sys

HERE = os.path.dirname(os.path.abspath(__file__))
PS1 = os.path.join(HERE, 'print-to-laser.ps1')
HTML = os.path.join(HERE, 'index.html')
OPEN = '<script type="text/plain" id="ps1src">'
CLOSE = '</script>'

ps1 = io.open(PS1, encoding='utf-8', newline='').read().replace('\r\n', '\n')
if '</script' in ps1.lower():
    sys.exit('The script now contains a </script sequence, which cannot be embedded as-is.')

html = io.open(HTML, encoding='utf-8', newline='').read()
i = html.find(OPEN)
if i < 0:
    sys.exit('No <script type="text/plain" id="ps1src"> block found in index.html.')
start = i + len(OPEN)
end = html.find(CLOSE, start)
if end < 0:
    sys.exit('The ps1src block in index.html is not closed.')

embedded = html[start:end]
wanted = '\n' + ps1                      # a leading newline HTML would eat anyway
# Git on Windows may hand either file back with CRLF, and that is not drift.
if embedded.replace('\r\n', '\n') == wanted:
    print('in sync (%d bytes)' % len(ps1))
    sys.exit(0)
if '--check' in sys.argv:
    sys.exit('OUT OF SYNC: index.html carries a different print-to-laser.ps1. '
             'Run python sync-bridge.py')
io.open(HTML, 'w', encoding='utf-8', newline='').write(html[:start] + wanted + html[end:])
print('updated index.html with print-to-laser.ps1 (%d bytes)' % len(ps1))
