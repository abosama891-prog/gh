@echo off
chcp 65001 >nul
title نشر التطبيق على Firebase
cls

echo ========================================
echo   نشر مراقبة سيارات شركة أبو غالي
echo ========================================
echo.

echo [1/2] جاري بناء المشروع...
echo ----------------------------------------
call npm run build
if errorlevel 1 (
    echo.
    echo ❌ فشل البناء! راجع الأخطاء فوق.
    echo.
    pause
    exit /b 1
)
echo.
echo ✅ تم البناء بنجاح
echo.

echo [2/2] جاري النشر على Firebase...
echo ----------------------------------------
call firebase deploy --only hosting
if errorlevel 1 (
    echo.
    echo ❌ فشل النشر! راجع الأخطاء فوق.
    echo.
    pause
    exit /b 1
)
echo.
echo ========================================
echo   ✅ تم النشر بنجاح!
echo   🌐 https://aboughaly.web.app
echo ========================================
echo.
pause