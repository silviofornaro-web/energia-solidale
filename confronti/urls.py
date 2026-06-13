from django.urls import path

from . import views


urlpatterns = [
    path("", views.confronto, name="confronto"),
    path("scarica-excel/", views.scarica_excel, name="scarica_excel"),
    path("area-clienti/confronto-illumia/", views.confronto_cliente_illumia, name="confronto_cliente_illumia"),
    path("area-clienti/scarica-excel/", views.scarica_excel_cliente_illumia, name="scarica_excel_cliente_illumia"),
]
