from django.urls import path

from . import views


urlpatterns = [
    path("", views.confronto, name="confronto"),
    path("scarica-excel/", views.scarica_excel, name="scarica_excel"),
]
