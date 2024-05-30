from time import sleep
import pyautogui
import click
import pygetwindow as gw

@click.command()
@click.argument("text", type=click.types.STRING)
@click.option("--wait", type=click.types.INT, default=1)
@click.option("--wnd", type=click.types.STRING)
@click.option("--interval", type=click.types.INT, default=0.1)
def main(text, wait, wnd, interval):
    if wnd:
        window : gw.Win32Window = gw.getWindowsWithTitle(wnd)[0]
        while True:
            if not window.isActive:
                sleep(1)
                continue
            break

    if wait:
        sleep(wait)

    pyautogui.typewrite(text, interval=interval)


if __name__ == "__main__":
    main()