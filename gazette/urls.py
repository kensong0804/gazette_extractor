# gazette/urls.py
from django.urls import path
from . import views

urlpatterns = [
    path("", views.upload_xml, name="upload_xml"),
]
