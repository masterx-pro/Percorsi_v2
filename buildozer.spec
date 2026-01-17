[app]
title = Percorsi Pro
package.name = percorsipro
package.domain = org.mattiaprosperi
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json
version = 2.0.0

# Requirements con supporto Excel
requirements = python3,kivy==2.2.1,requests,urllib3,certifi,charset-normalizer,idna,pillow,openpyxl

orientation = portrait
fullscreen = 0
android.presplash_color = #FF1DA0
android.permissions = INTERNET,ACCESS_NETWORK_STATE,WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE,ACCESS_FINE_LOCATION
android.api = 33
android.minapi = 21
android.ndk_api = 21
android.accept_sdk_license = True
android.archs = arm64-v8a
log_level = 2
warn_on_root = 0
