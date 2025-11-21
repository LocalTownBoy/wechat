"""wxcloudrun URL Configuration

The `urlpatterns` list routes URLs to views. For more information please see:
    https://docs.djangoproject.com/en/3.2/topics/http/urls/
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

from wxcloudrun import views
from django.conf.urls import url

urlpatterns = (
    # 推送消息写入Excel
    url(r'^^push/msg(/)?$', views.push_msg),
    # 展示微信公众号消息列表
    url(r'^^push/list(/)?$', views.push_msg_list),
    # 论文收集页面
    url(r'^^papers(/)?$', views.papers),
    # 主动向公众号用户推送文本（需要配置 WX_APPID / WX_SECRET）
    url(r'^^wx/send(/)?$', views.wx_send_message),
    # 批量获取关注用户信息
    url(r'^^wx/users/info(/)?$', views.wx_users_info),

    # 计数器接口
    url(r'^^api/count(/)?$', views.counter),

    # 获取主页
    url(r'(/)?$', views.index),
)
