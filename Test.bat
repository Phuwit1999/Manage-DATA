@echo off
title DATA System
:: ย้ายตำแหน่งไปยังโฟลเดอร์ที่เก็บไฟล์
cd /d "C:\Users\ITSuperSupport\Desktop\CH"
:: สั่งรัน streamlit ผ่าน python module
python -m streamlit run TESTDATA.py
pause
