from pywinauto import Desktop

print("=" * 80)

for w in Desktop(backend="uia").windows():

    try:

        print("TITLE :", repr(w.window_text()))
        print("CLASS :", w.class_name())
        print("-" * 80)

    except:
        pass