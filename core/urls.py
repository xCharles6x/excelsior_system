from django.urls import path
from . import views

urlpatterns = [
    # Auth
    path('login/',    views.login_view,   name='login'),
    path('logout/',   views.logout_view,  name='logout'),

    # Dashboard
    path('dashboard/', views.dashboard, name='dashboard'),

    # Records (Create)
    path('loadtest/',   views.loadtest,   name='loadtest'),
    path('inventory/',  views.inventory,  name='inventory'),
    path('repair/',     views.repair,     name='repair'),
    path('inspection/', views.inspection, name='inspection'),

    # Records (List & Export)
    path('list/',   views.list_records, name='list_records'),
    path('export/', views.export_excel, name='export_excel'),

    # Records (Edit)
    path('edit/loadtest/<int:pk>/',  views.edit_loadtest,  name='edit_loadtest'),
    path('edit/inventory/<int:pk>/', views.edit_inventory, name='edit_inventory'),
    path('edit/repair/<int:pk>/',    views.edit_repair,    name='edit_repair'),

    # Records (Delete)
    path('delete/loadtest/<int:pk>/',  views.delete_loadtest,  name='delete_loadtest'),
    path('delete/inventory/<int:pk>/', views.delete_inventory, name='delete_inventory'),
    path('delete/repair/<int:pk>/',    views.delete_repair,    name='delete_repair'),

    # Download Excel Forms
    path('download/repair/<int:pk>/',     views.download_repair_excel,     name='download_repair_excel'),
    path('download/inventory/<int:pk>/',  views.download_inventory_excel,  name='download_inventory_excel'),

    # Certificates
    path('certificates/',                    views.certificate_list,     name='certificate_list'),
    path('certificates/new/',               views.certificate_create,   name='certificate_create'),
    path('certificates/<int:pk>/edit/',     views.certificate_edit,     name='certificate_edit'),
    path('certificates/<int:pk>/download/', views.certificate_download, name='certificate_download'),
    path('certificates/<int:pk>/delete/',   views.certificate_delete,   name='certificate_delete'),

    # Customers
    path('customers/new/',             views.customer_create,      name='customer_create'),
    path('customers/<int:pk>/delete/', views.customer_delete,      name='customer_delete'),
    path('api/customer-address/',      views.customer_address_api, name='customer_address_api'),
    path('api/loadcell-cert/save/',    views.loadcell_cert_save,   name='loadcell_cert_save'),
    path('api/loadcell-cert/delete/',  views.loadcell_cert_delete, name='loadcell_cert_delete'),

    # Autocomplete API
    path('api/autocomplete/',          views.autocomplete_api,          name='autocomplete_api'),
    path('api/next-cert-number/',      views.next_certificate_number,   name='next_certificate_number'),

    # Go to Certification (pre-fill from records)
    path('cert-from/loadtest/<int:pk>/',  views.cert_from_loadtest,  name='cert_from_loadtest'),
    path('cert-from/repair/<int:pk>/',    views.cert_from_repair,    name='cert_from_repair'),
    path('cert-from/inventory/<int:pk>/', views.cert_from_inventory, name='cert_from_inventory'),

    # Admin: User / Account Management
    path('admin-panel/users/',                 views.admin_user_list,   name='admin_user_list'),
    path('admin-panel/users/create/',          views.admin_user_create, name='admin_user_create'),
    path('admin-panel/users/<int:pk>/edit/',   views.admin_user_edit,   name='admin_user_edit'),
    path('admin-panel/users/<int:pk>/delete/', views.admin_user_delete, name='admin_user_delete'),

    # Photo Reference
    path('photos/',                  views.photo_gallery, name='photo_gallery'),
    path('photos/upload/',           views.photo_upload,  name='photo_upload'),
    path('photos/<int:pk>/detail/',  views.photo_detail,  name='photo_detail'),
    path('photos/<int:pk>/delete/',  views.photo_delete,  name='photo_delete'),
]