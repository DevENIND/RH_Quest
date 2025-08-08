import flet as ft

def main(page: ft.Page):
    page.title = "App Flet Locaweb"
    page.add(ft.Text("Tela de login funcionando!"))

if __name__ == "__main__":
    ft.app(
        target=main,
        view=ft.WEB_BROWSER,
        host="0.0.0.0",
        port=8000
    )
