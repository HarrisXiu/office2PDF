"""Local smoke checks; --engine office/wps also exercises installed applications."""
import argparse
import os
from pathlib import Path
import tempfile

from office2pdf_v5 import PDFUltimateApp, AppConfig, TkinterDnD


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--engine', choices=('office', 'wps'))
    args = parser.parse_args()
    fixtures = Path(__file__).parent / 'test_artifacts'
    with tempfile.TemporaryDirectory() as work:
        os.chdir(work)
        root = None
        try:
            for scale in (1.0, 1.25, 1.5, 2.0):
                if root is not None:
                    for callback in root.tk.call('after', 'info'):
                        root.after_cancel(callback)
                    root.destroy()
                root = TkinterDnD.Tk()
                root.withdraw()
                root.tk.call('tk', 'scaling', 96 / 72 * scale)
                app = PDFUltimateApp(root)
                for width, height in ((1000, 740), (800, 600), (560, 420)):
                    root.geometry(f'{width}x{height}')
                    root.deiconify()
                    root.update()
                    for show_log in (False, True):
                        if app.show_log_var.get() != show_log:
                            app.log_toggle.invoke()
                        # Check fixed controls with the settings scrolled to either end.
                        for fraction in (0.0, 1.0):
                            app.settings_canvas.yview_moveto(fraction)
                            root.update_idletasks()
                            for button in (app.btn_convert, app.btn_cancel):
                                x = button.winfo_rootx() - root.winfo_rootx()
                                y = button.winfo_rooty() - root.winfo_rooty()
                                assert x >= 0 and y >= 0
                                assert x + button.winfo_width() <= root.winfo_width()
                                assert y + button.winfo_height() <= root.winfo_height()
            print('PASS: fixed action buttons at 3 window sizes / 4 Tk scales', flush=True)
            if args.engine:
                cfg = AppConfig(engine=args.engine)
                app.config.engine = args.engine
                for kind, suffix in (('Word', 'docx'), ('PowerPoint', 'pptx')):
                    source = {'path': str(fixtures / f'sample.{suffix}'), 'range': ''}
                    out = os.path.join(work, f'{kind}.pdf')
                    assert app.cv_document(kind, source, out, cfg), kind
                    print(f'PASS: {args.engine} {kind} export', flush=True)
                source = {'path': str(fixtures / 'sample.xlsx'), 'range': ''}
                assert app.get_excel_sheets(source['path']) == ['First', 'Second']
                first = app.cv_excel_units(source, work, cfg)
                second = app.cv_excel_units(source, work, cfg)
                assert len(first) == len(second) == 2
                assert not {p for p, _ in first} & {p for p, _ in second}
                print(f'PASS: {args.engine} Excel sheets, export, unique output paths', flush=True)
        finally:
            while not app.progress_queue.empty():
                print(app.progress_queue.get(), flush=True)
            root.destroy()
            os.chdir(fixtures.parent)


if __name__ == '__main__':
    main()
