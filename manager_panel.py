import flet as ft

def show_manager_panel(page):
    def close_dialog(e):
        page.dialog.open = False
        page.update()

    def show_message(title, message):
        dlg = ft.AlertDialog(
            modal=True,
            title=ft.Text(title),
            content=ft.Text(message),
            actions=[ft.TextButton("باشه", on_click=close_dialog)]
        )
        page.dialog = dlg
        dlg.open = True
        page.update()

    page.clean()
    page.add(
        ft.Text("پنل مدیر", size=24, weight=ft.FontWeight.BOLD),
        ft.ElevatedButton("خروج", on_click=lambda e: show_message("خروج", "موفقیت‌آمیز"))
    )
    page.update()