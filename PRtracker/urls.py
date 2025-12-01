from django.contrib import admin
from django.urls import path
from core import views

urlpatterns = [
    path("admin/", admin.site.urls),

    # 메인: step1~3 한 페이지
    path("", views.home, name="home"),

    # 예전 /copy/ 주소로 들어와도 같은 페이지로
    path("copy/", views.home, name="copy_only"),
]
