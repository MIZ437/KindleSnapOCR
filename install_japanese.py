"""日本語言語パックインストールスクリプト"""
import sys
sys.path.insert(0, 'src')
from tesseract_installer import ensure_japanese_installed, is_language_installed

def progress(status):
    print('  ' + status)

print('Checking Japanese language pack...')
print('')

if is_language_installed('jpn'):
    print('  Japanese (jpn) is already installed')
else:
    print('  Japanese (jpn) not found, downloading...')

if is_language_installed('jpn_vert'):
    print('  Japanese vertical (jpn_vert) is already installed')
else:
    print('  Japanese vertical (jpn_vert) not found, downloading...')

print('')
success = ensure_japanese_installed(progress)

print('')
if success:
    print('Done! Japanese language pack is ready.')
else:
    print('Failed to install Japanese language pack.')

input('Press Enter to close...')
