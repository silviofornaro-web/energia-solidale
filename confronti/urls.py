from django.urls import path

from . import views


urlpatterns = [
    path("", views.confronto, name="confronto"),
    path("area-clienti/", views.accesso_clienti, name="accesso_clienti"),
    path("scarica-excel/", views.scarica_excel, name="scarica_excel"),
    path("archivio-report/", views.archivio_report, name="archivio_report"),
    path("archivio-report/salva/", views.archivia_report_corrente, name="archivia_report_corrente"),
    path(
        "archivio-report/cartella/<int:folder_id>/",
        views.archivio_report_cartella,
        name="archivio_report_cartella",
    ),
    path(
        "archivio-report/report/<int:report_id>/scarica/",
        views.scarica_report_archiviato,
        name="scarica_report_archiviato",
    ),
    path("area-clienti/confronto-illumia/", views.confronto_cliente_illumia, name="confronto_cliente_illumia"),
    path("area-clienti/scarica-excel/", views.scarica_excel_cliente_illumia, name="scarica_excel_cliente_illumia"),
]
