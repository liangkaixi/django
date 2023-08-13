from django.urls import path
from . import views

urlpatterns = [
    path('', views.class_allocation_form, name='class_allocation_form'),
    path('allocate/', views.class_allocation, name='class_allocation'),
]
