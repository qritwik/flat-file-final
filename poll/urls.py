from django.conf.urls import url
from poll import views

app_name = 'poll'

urlpatterns = [
    url(r'^$',views.index,name="index"),
    url(r'^final/',views.final,name="final"),

]
