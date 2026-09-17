from bs4 import BeautifulSoup

with open('tc14_export_dialog_dom.html', 'r', encoding='utf-8') as f:
    soup = BeautifulSoup(f.read(), 'html.parser')

dialog = soup.find(class_=lambda c: c and 'sw-right-dialog' in c)
if dialog:
    print('=== Export Dialog Elements ===')
    for elem in dialog.find_all(['button', 'input', 'label', 'a']):
        txt = elem.text.strip().replace('\n', ' ')
        classes = ' '.join(elem.get('class', []))
        if elem.name in ['button', 'input']:
            print(f"[{elem.name}] type='{elem.get('type')}', id='{elem.get('id')}', name='{elem.get('name')}', action='{elem.get('action')}', text='{txt}', class='{classes}'")
        elif elem.name == 'label' and txt:
            print(f"[label] for='{elem.get('for')}', text='{txt}'")
else:
    print('Dialog not found.')
