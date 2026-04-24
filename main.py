import webview
import os
import sys

class Api:
    def __init__(self):
        self.window = None

    def save_csv(self, filename, content):
        if self.window is None:
            return False
            
        result = self.window.create_file_dialog(
            webview.SAVE_DIALOG, 
            directory=os.environ.get('HOME', os.path.expanduser('~')), 
            save_filename=filename
        )
        if result:
            filepath = result[0] if isinstance(result, (tuple, list)) else result
            try:
                with open(filepath, 'w', encoding='utf-8-sig') as f:
                    f.write(content)
                return True
            except Exception as e:
                print(f"Failed to save file: {e}")
                return False
        return False

def create_app():
    api = Api()
    html_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'index.html')
    
    window = webview.create_window(
        'Vanning Layout Simulator', 
        url=f'file://{html_path}',
        width=1380, 
        height=850,
        resizable=True,
        min_size=(1024, 768),
        js_api=api
    )
    api.window = window
    
    webview.start(debug=False)

if __name__ == '__main__':
    create_app()
