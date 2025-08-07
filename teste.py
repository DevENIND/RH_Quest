import flet as ft
import base64

def main(page: ft.Page):
    with open("Imagem_Quest.png", "rb") as img_file:
        img_base64 = base64.b64encode(img_file.read()).decode('utf-8')

    imagem_fundo = ft.Image(
        src=f"data:image/png;base64,{img_base64}",
        fit=ft.ImageFit.COVER,
        expand=True
    )

    conteudo = ft.Text("Conteúdo sobre a imagem", color=ft.Colors.BLACK)

    page.add(ft.Stack([
        imagem_fundo,
        conteudo
    ]))

ft.app(target=main, view=ft.WEB_BROWSER)
