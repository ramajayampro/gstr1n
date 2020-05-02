from django.contrib import admin
from django.urls import path, include
from todolist_app import views as todolist_views



urlpatterns = [
    path('admin/', admin.site.urls),
    path('', todolist_views.index1a, name='index1a'),
    path('zip/', todolist_views.index2, name='index2'),
    path('todolist/', include('todolist_app.urls')),
    path('contact', todolist_views.contact, name='contact'),
    path('about-us', todolist_views.about, name='about'),
    path('mul/', todolist_views.FileFieldView, name='FileFieldView')
    
    ]
