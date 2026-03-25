from django.contrib import admin
from django.urls import path, include
from django.shortcuts import redirect

# optional: home view to redirect based on authentication
def home(request):
    if request.user.is_authenticated:
        return redirect('dashboard')
    return redirect('login')

urlpatterns = [
    path('admin/', admin.site.urls),
    path('', home, name='home'),         # root redirect
    path('', include('core.urls')),      # include all app URLs
]