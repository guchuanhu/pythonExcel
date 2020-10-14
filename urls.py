"""myExcel URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/3.0/topics/http/urls/
Examples:
Function views
    1. Add an import:  from my_app import views
    2. Add a URL to urlpatterns:  path('', views.home, name='home')
Class-based views
    1. Add an import:  from other_app.views import Home
    2. Add a URL to urlpatterns:  path('', Home.as_view(), name='home')
Including another URLconf
    1. Import the include() function: from django.urls import include, path
    2. Add a URL to urlpatterns:  path('blog/', include('blog.urls'))
"""
from django.conf.urls import url
from django.views.generic.base import TemplateView
 
from . import views,search
 
urlpatterns = [
    # url(r'^$', TemplateView.as_view(template_name="index.html")),
    url(r'^$', search.search_form),
    url(r'^search$', search.search),
    url(r'^searchDir$', search.searchDir),
    url(r'^getOthersApi$', search.getOthersApi),
    url(r'^searchTest$', search.searchTest),
    url(r'^tamaker$', search.tamaker),
]