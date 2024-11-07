from django.urls import path
from . import views
# from .views import import_from_excel

urlpatterns = [
#    path('',views.import_from_excel, name='import_from_excel'),
    path('', views.home, name='home'),
    path('index/', views.index, name='index'),
    path('geocalc/', views.geocalc, name='geocalc'),
    path('report/', views.report, name='report'),
    path('report1/', views.report, name='report1'),
    path('method/', views.method, name='method'),
    path('contacts/', views.contacts, name='contacts'),
    path('help/', views.help, name='help'),
    path('map/', views.map, name='map'),
]
