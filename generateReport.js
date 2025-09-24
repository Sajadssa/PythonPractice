(function(){
    "use strict";

    var busy = false;
    var debounceTimer = null;
    var DEBUG = true;

    // ==================== توابع پایه ====================
    function log(message, data) {
        if (DEBUG) {
            try {
                if (data === undefined) console.log("[شماره‌گیر] " + message);
                else console.log("[شماره‌گیر] " + message, data);
            } catch(e){}
        }
    }

    function getControlValue(name) {
        try {
            var ctrl = form.GetControl(name);
            if (!ctrl) return "";
            if (typeof ctrl.GetValue === "function") return (ctrl.GetValue() || "").toString().trim();
            if (typeof ctrl.getValue === "function") return (ctrl.getValue() || "").toString().trim();
            if (ctrl.value !== undefined) return (ctrl.value || "").toString().trim();
            if (ctrl.Value !== undefined) return (ctrl.Value || "").toString().trim();
            return "";
        } catch (e) {
            log("خطا در خواندن " + name, e.message || e);
            return "";
        }
    }

    function setControlValue(name, value) {
        try {
            var ctrl = form.GetControl(name);
            if (!ctrl) return false;
            if (typeof ctrl.SetValue === "function") { ctrl.SetValue(value); }
            else if (typeof ctrl.setValue === "function") { ctrl.setValue(value); }
            else if (ctrl.value !== undefined) { ctrl.value = value; }
            else if (ctrl.Value !== undefined) { ctrl.Value = value; }
            try { if (typeof ctrl.Refresh === "function") ctrl.Refresh(); } catch(e){}
            return true;
        } catch (e) {
            log("خطا در تنظیم " + name, e.message || e);
            return false;
        }
    }

    function pad4(num) {
        var str = (num === null || num === undefined) ? "" : num.toString();
        while (str.length < 4) str = "0" + str;
        return str;
    }

    // ==================== ساخت Pattern ها ====================
    
    // BasePattern بدون LocationCode - فقط برای محاسبه شماره
    function createBasePattern() {
        var constantPart = getControlValue("c_ConstantPart");
        var contCode = getControlValue("c_ContCode");
        var psCode = getControlValue("c_PSCode");
        var spCode = getControlValue("c_SPCode");
        var maingrCode = getControlValue("c_MaingrCode");
        var tpCode = getControlValue("c_TPCode");
        
        var basePattern = constantPart + contCode + "-" + psCode + spCode + "-" + maingrCode + tpCode;
        log("BasePattern (بدون LocationCode):", basePattern);
        return basePattern;
    }

    // FullPattern با LocationCode - فقط برای نمایش نهایی
    function createFullPattern() {
        var constantPart = getControlValue("c_ConstantPart");
        var locationCode = getControlValue("c_LocationCode");
        var contCode = getControlValue("c_ContCode");
        var psCode = getControlValue("c_PSCode");
        var spCode = getControlValue("c_SPCode");
        var maingrCode = getControlValue("c_MaingrCode");
        var tpCode = getControlValue("c_TPCode");
        
        var fullPattern = constantPart + locationCode + contCode + "-" + psCode + spCode + "-" + maingrCode + tpCode;
        log("FullPattern (برای نمایش):", fullPattern);
        return fullPattern;
    }

    // ==================== خواندن حداکثر شماره از کنترل‌ها ====================
    function getMaxFromControl() {
        log("خواندن حداکثر شماره از c_MaxRep");
        
        var maxRep = getControlValue("c_MaxRep").trim();
        if (!maxRep) {
            log("c_MaxRep خالی است");
            return null;
        }
        
        log("c_MaxRep:", maxRep);
        
        // استخراج شماره از فرمت: PATTERN-0516-G00
        var match = maxRep.match(/-(\d+)-[A-Za-z0-9]+$/);
        if (match) {
            var maxNumber = parseInt(match[1], 10);
            log("حداکثر شماره استخراج شده:", maxNumber);
            return maxNumber;
        }
        
        log("فرمت c_MaxRep نامعتبر");
        return null;
    }

    // ==================== خواندن DefaultReport ====================
    function getDefaultFromControl() {
        log("خواندن DefaultReport از c_DefRep");
        
        var defRep = getControlValue("c_DefRep").trim();
        if (!defRep) {
            log("c_DefRep خالی است");
            return null;
        }
        
        log("c_DefRep:", defRep);
        
        // اگر فقط عدد باشد
        var asNum = parseInt(defRep, 10);
        if (!isNaN(asNum) && /^\d+$/.test(defRep)) {
            log("DefaultReport عددی:", asNum);
            return asNum;
        }
        
        // اگر فرمت کامل باشد: PATTERN-0459-G00
        var match = defRep.match(/-(\d+)-[A-Za-z0-9]+$/);
        if (match) {
            var defaultNumber = parseInt(match[1], 10);
            log("DefaultReport از فرمت کامل:", defaultNumber);
            return defaultNumber;
        }
        
        log("فرمت c_DefRep نامعتبر");
        return null;
    }

    // ==================== محاسبه شماره بعدی ====================
    function calculateNextNumber() {
        log("=== محاسبه شماره بعدی ===");
        
        var maxNumber = getMaxFromControl();
        var defaultNumber = getDefaultFromControl();
        
        log("Max از کنترل:", maxNumber, "Default از کنترل:", defaultNumber);
        
        var nextNumber = null;
        
        if (maxNumber !== null) {
            // اگر Max موجود است، شماره بعدی = Max + 1
            nextNumber = maxNumber + 1;
            log("شماره بعدی = Max + 1 =", nextNumber);
        } else if (defaultNumber !== null) {
            // اگر Max نیست اما Default هست، از Default استفاده کن
            nextNumber = defaultNumber;
            log("شماره بعدی = Default =", nextNumber);
        } else {
            // اگر هیچکدام نیست، از 1 شروع کن
            nextNumber = 1;
            log("شماره بعدی = پیش‌فرض =", nextNumber);
        }
        
        log("=== شماره نهایی محاسبه شده:", nextNumber, "===");
        return nextNumber;
    }

    // ==================== آپدیت Min/Max ====================
    function updateMinMax(newNumber) {
        log("آپدیت Min/Max برای شماره", newNumber);
        
        var fullPattern = createFullPattern();
        var newReportNo = fullPattern + "-" + pad4(newNumber) + "-G00";
        
        // آپدیت Max
        var currentMax = getControlValue("c_MaxRep").trim();
        var shouldUpdateMax = true;
        
        if (currentMax) {
            var maxMatch = currentMax.match(/-(\d+)-/);
            if (maxMatch) {
                var currentMaxNum = parseInt(maxMatch[1], 10);
                shouldUpdateMax = (newNumber > currentMaxNum);
            }
        }
        
        if (shouldUpdateMax) {
            setControlValue("c_MaxRep", newReportNo);
            log("c_MaxRep آپدیت شد:", newReportNo);
        }
        
        // آپدیت Min
        var currentMin = getControlValue("c_MinRep").trim();
        var shouldUpdateMin = true;
        
        if (currentMin) {
            var minMatch = currentMin.match(/-(\d+)-/);
            if (minMatch) {
                var currentMinNum = parseInt(minMatch[1], 10);
                shouldUpdateMin = (newNumber < currentMinNum);
            }
        }
        
        if (shouldUpdateMin) {
            setControlValue("c_MinRep", newReportNo);
            log("c_MinRep آپدیت شد:", newReportNo);
        }
    }

    // ==================== اعتبارسنجی ورودی‌ها ====================
    function validateInputs() {
        var checks = [
            { field: "c_PSCode", message: "لطفاً پروسس را انتخاب کنید" },
            { field: "c_SPCode", message: "لطفاً زیر پروسس را انتخاب کنید" },
            { field: "c_MaingrCode", message: "لطفاً گروه اصلی را انتخاب کنید" },
            { field: "c_TPCode", message: "لطفاً بازه گزارش را انتخاب کنید" },
            { field: "c_Subject", message: "لطفاً عنوان گزارش را انتخاب کنید" },
            { field: "c_LocationCode", message: "لطفاً موقعیت را انتخاب کنید" },
            { field: "c_ContCode", message: "لطفاً پیمانکار را انتخاب کنید" },
            { field: "c_ReportDate", message: "لطفاً تاریخ گزارش را انتخاب کنید" }
        ];
        
        for (var i = 0; i < checks.length; i++) {
            if (!getControlValue(checks[i].field)) {
                return checks[i].message;
            }
        }
        return null;
    }

    // ==================== تولید شماره گزارش ====================
    function generateReportNumber(showAlerts) {
        log("=== شروع تولید شماره گزارش ===");
        
        if (busy) {
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(function() { generateReportNumber(showAlerts); }, 50);
            return;
        }
        
        // بررسی اینکه آیا قبلاً شماره تولید شده یا نه
        var existingReportNo = getControlValue("c_ReportNo").trim();
        if (existingReportNo && showAlerts) {
            log("شماره گزارش قبلاً تولید شده:", existingReportNo);
            alert("شماره گزارش قبلاً تولید شده است: " + existingReportNo);
            return false;
        }


                busy = true;
        
        try {
            // اعتبارسنجی
            if (showAlerts) {
                var validationError = validateInputs();
                if (validationError) {
                    alert(validationError);
                    busy = false;
                    return false;
                }
            }
            
            var basePattern = createBasePattern();
            var fullPattern = createFullPattern();
            
            // ذخیره BasePattern در c_pattern
            setControlValue("c_pattern", basePattern);
            
            // اگر در حالت silent validation نیاز باشد
            if (!showAlerts && validateInputs()) {
                busy = false;
                return false;
            }
            
            // محاسبه شماره بعدی
            var nextNumber = calculateNextNumber();
            
            // آپدیت Min/Max
            updateMinMax(nextNumber);
            
            // ساخت شماره نهایی
            var paddedNumber = pad4(nextNumber);
            var finalReportNo = fullPattern + "-" + paddedNumber + "-G00";
            
            // تنظیم کنترل‌ها
            setControlValue("c_PartNum", paddedNumber);
            setControlValue("c_ReportNo", finalReportNo);
            setControlValue("c_Rev", "G00");
            
            // غیرفعال کردن کنترل ReportNo
            try {
                var reportControl = form.GetControl("c_ReportNo");
                if (reportControl) {
                    if (typeof reportControl.SetEnabled === "function") {
                        reportControl.SetEnabled(false);
                    } else if (reportControl.disabled !== undefined) {
                        reportControl.disabled = true;
                    }
                }
            } catch(e) {}
            
            // ذخیره فرم
            try {
                if (typeof form.Save === "function") {
                    form.Save();
                    log("فرم ذخیره شد");
                }
            } catch(e) {
                log("خطا در ذخیره فرم:", e);
            }
            
            log("=== تولید شماره تکمیل شد:", finalReportNo, "===");
            busy = false;
            return true;
            
        } catch(e) {
            log("خطای کلی در تولید شماره:", e);
            busy = false;
            return false;
        }
    }

    // ==================== اتصال رویدادها ====================
    function attachEvents() {
        log("اتصال رویدادها");
        
        var controlNames = ["c_ConstantPart", "c_LocationCode", "c_MaingrCode", "c_ContCode", 
                           "c_PSCode", "c_SPCode", "c_TPCode", "c_Subject", "c_ReportDate", 
                           "c_MaxRep", "c_DefRep"];
                           
        var handler = function() {
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(function() { generateReportNumber(false); }, 100);
        };
        
        controlNames.forEach(function(name) {
            try {
                var ctrl = form.GetControl(name);
                if (ctrl) {
                    if (ctrl.SelectionChanged && typeof ctrl.SelectionChanged.connect === "function") {
                        ctrl.SelectionChanged.connect(handler);
                    } else if (typeof ctrl.onchange !== "undefined") {
                        var oldHandler = ctrl.onchange;
                        ctrl.onchange = function(e) {
                            try {
                                if (oldHandler) oldHandler(e);
                            } catch(ex) {}
                            handler();
                        };
                    }
                }
            } catch(e) {
                log("خطا در اتصال رویداد برای " + name + ":", e);
            }
        });
        
        // اتصال دکمه Generate
        try {
            var btn = form.GetControl("c_control1");
            if (btn) {
                if (btn.Click && typeof btn.Click.connect === "function") {
                    btn.Click.connect(function() { generateReportNumber(true); });
                } else if (typeof btn.onclick !== "undefined") {
                    btn.onclick = function() { generateReportNumber(true); };
                }
            }
        } catch(e) {
            log("خطا در اتصال دکمه:", e);
        }
    }

    // ==================== راه‌اندازی ====================
    function initialize() {
        log("=== راه‌اندازی شماره‌گیر ===");
        try {
            attachEvents();
            setTimeout(function() { generateReportNumber(false); }, 200);
            window.generateReportNumber = generateReportNumber;
            log("=== راه‌اندازی تکمیل شد ===");
        } catch(e) {
            log("خطا در راه‌اندازی:", e);
        }
    }

    initialize();

})();