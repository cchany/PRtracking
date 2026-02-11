# core/urls.py
from django.urls import path
from . import views

urlpatterns = [
    path("", views.home, name="home"),
    path("news/collect/", views.news_collect, name="news_collect"),
    path("news/download/<str:job_id>/", views.news_download, name="news_download"),
]
