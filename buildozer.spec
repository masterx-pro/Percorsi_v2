[app]
title = Percorsi Pro
package.name = percorsipro
package.domain = org.mattiaprosperi
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json
version = 3.0.0

# Requirements SENZA pandas (causa crash su Android)
requirements = python3,kivy==2.2.1,requests,urllib3,certifi,charset-normalizer,idna,openpyxl,et_xmlfile

orientation = portrait
fullscreen = 0
android.presplash_color = #FF1DA0
android.permissions = INTERNET,ACCESS_NETWORK_STATE,WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE
android.api = 33
android.minapi = 21
android.ndk_api = 21
android.accept_sdk_license = True
android.archs = arm64-v8a
android.allow_backup = True

# Importante per evitare crash
android.enable_androidx = True
p4a.branch = master

log_level = 2
warn_on_root = 0
