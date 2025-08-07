import flet as ft
import asyncio
import time

IDLE_TIMEOUT = 60  # segundos

def main(page: ft.Page):

    
    page.add(
        ft.ElevatedButton(
            "Abrir PDF no navegador",
            on_click=lambda e: page.launch_url("https://enindengenharia.sharepoint.com/:b:/s/TesteSiteInterno/Ed9WO1N9KplHkD5g14-sGdgBEmFhlYzcSEbyrz4kW0LoVw?e=99i7OG")
        )
    )

ft.app(target=main)


