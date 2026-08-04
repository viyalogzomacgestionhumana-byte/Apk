[app]
title = Buscador GH
package.name = buscadorgh
package.domain = org.test
source.dir = .
source.include_exts = py,png,jpg,kv,atlas
version = 0.1

requirements = python3==3.11.6,hostpython3==3.11.6,kivy==2.3.0,kivymd==1.2.0,openpyxl,requests,urllib3,python-dateutil

orientation = portrait
fullscreen = 0

android.permissions = INTERNET
android.api = 33
android.minapi = 24
android.ndk = 25b
android.archs = arm64-v8a, armeabi-v7a
android.allow_backup = True
android.accept_sdk_license = True

p4a.branch = develop

[buildozer]
log_level = 2
warn_on_root = 1
