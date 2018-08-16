from django.urls import path

from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path(r'takko', views.upload_file, name='upload_file'),
    path(r'invoice', views.invoice_test, name='invoice_test'),
]