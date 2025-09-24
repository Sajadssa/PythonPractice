(function(){
    "use strict";

    var busy = false;
    var debounceTimer = null;
    var DEBUG = true;

    // ==================== بخش 1: توابع پایه ====================
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

    // ==================== بخش 2: خواندن DefaultReport ====================
    function getDefaultReportFromControl(basePattern) {
        try {
            var defRep = getControlValue("c_DefRep") || "";
            if (!defRep) return null;
            // اگر فقط عدد
            var asNum = parseInt(defRep,10);
            if (!isNaN(asNum) && /^\d+$/.test(defRep)) return asNum;
            var m = defRep.match(/^(.+)-(\d+)(?:-[A-Za-z0-9]+)?$/);
            if (m) {
                var defBase = m[1].trim();
                var defNum = parseInt(m[2],10);
                if (defBase === basePattern) return defNum;
                var locationCode = getControlValue("c_LocationCode").trim();
                if (locationCode && defBase.indexOf(locationCode) !== -1) {
                    var withoutLoc = defBase.split(locationCode).join("").trim();
                    if (withoutLoc === basePattern) return defNum;
                }
            }
        } catch(e){ log("getDefaultReportFromControl error", e); }
        return null;
    }

    function getDefaultReportFromList(basePattern) {
        try {
            var defaultReports = [];
            if (form.DefaultReport && Array.isArray(form.DefaultReport)) defaultReports = form.DefaultReport;
            else if (typeof form.GetDefaultReportList === "function") {
                try { var res = form.GetDefaultReportList(); if (Array.isArray(res)) defaultReports = res; } catch(e){}
            }
            for (var i=0;i<defaultReports.length;i++){
                var it = defaultReports[i];
                if (!it) continue;
                // اگر آبجکت با فیلدهاست
                var bp = null, dr = null;
                try {
                    if (typeof it.get_item === "function") {
                        bp = it.get_item("BasePattern") || it.get_item("Pattern") || it.get_item("Title");
                        dr = it.get_item("DefaultReport") || it.get_item("ReportNo");
                    } else if (it.BasePattern || it.Pattern || it.Title) {
                        bp = it.BasePattern || it.Pattern || it.Title;
                        dr = it.DefaultReport || it.ReportNo || it["DefaultReport"];
                    } else {
                        // fallback به رشته
                        var s = it.toString();
                        var m = s.match(/^(.+)-(\d+)(?:-[A-Za-z0-9]+)?$/);
                        if (m) { bp = m[1]; dr = m[2]; }
                    }
                } catch(e){}
                if (!bp) continue;
                bp = bp.toString().trim();
                if (bp !== basePattern) continue;
                if (dr) {
                    var n = parseInt(dr.toString().trim(),10);
                    if (!isNaN(n)) return n;
                }
            }
        } catch(e){ log("getDefaultReportFromList error", e); }
        return null;
    }

    // ==================== بخش 3: خواندن DocumentEntryList به صورت مقاوم ====================
    function tryExtractReportNoFromEntry(entry) {
        // entry ممکنه string، آبجکت SP ListItem، یا آبجکت دیگر باشه
        try {
            if (!entry) return null;
            // اگر رشته هست
            if (typeof entry === "string") {
                var s = entry.trim();
                if (s.length === 0) return null;
                return s;
            }
            // اگر آبجکت با get_item (SP list item)
            if (typeof entry.get_item === "function") {
                var candidates = ["ReportNo","Report_No","Report","Title"];
                for (var i=0;i<candidates.length;i++){
                    try {
                        var v = entry.get_item(candidates[i]);
                        if (v) return v.toString().trim();
                    } catch(e){}
                }
            }
            // اگر آبجکت ساده با فیلدها
            var keys = ["ReportNo","Report_No","Report","Title","reportNo","report"];
            for (var k=0;k<keys.length;k++){
                if (entry[keys[k]]) {
                    try { return entry[keys[k]].toString().trim(); } catch(e){}
                }
            }
            // fallback: stringify and try to find pattern
            try {
                var js = JSON.stringify(entry);
                if (js && js.length) {
                    // تلاش برای پیدا کردن توکن‌هایی شبیه to ReportNo داخل json
                    var m = js.match(/([A-Za-z0-9\-]+-\d{1,4}-[A-Za-z0-9]+)/);
                    if (m && m[1]) return m[1];
                }
            } catch(e){}
        } catch(e){ log("tryExtractReportNoFromEntry error", e); }
        return null;
    }

    function getDocumentList() {
        // برمی‌گرداند آرایه‌ای از رشته‌های ReportNo (مثلاً "ABC-0001-G00")
        log("getDocumentList: شروع جمع‌آوری اسناد");
        var docs = [];
        try {
            // تلاش 1: اگر DocumentEntryList یک آرایه ساده دارد
            if (form.DocumentEntryList && Array.isArray(form.DocumentEntryList)) {
                log("getDocumentList: form.DocumentEntryList موجود، طول:", form.DocumentEntryList.length);
                form.DocumentEntryList.forEach(function(entry){
                    var r = tryExtractReportNoFromEntry(entry);
                    if (r) docs.push(r);
                });
            }
        } catch(e){ log("getDocumentList step1 error", e); }

        // تلاش 2: اگر متدهای کمکی وجود دارند
        var tryNames = ["GetDocumentEntryList","GetDocumentEntries","GetDocumentList","DocumentEntries","DocumentList"];
        for (var i=0;i<tryNames.length;i++){
            try {
                var fn = form[tryNames[i]];
                if (!fn) continue;
                if (typeof fn === "function") {
                    try {
                        var res = fn();
                        if (Array.isArray(res)) {
                            res.forEach(function(entry){
                                var r = tryExtractReportNoFromEntry(entry);
                                if (r) docs.push(r);
                            });
                            log("getDocumentList: خواندن از متد", tryNames[i], "تعداد:", res.length);
                        } else if (res && typeof res === "object") {
                            // برخی API ها آرایه-like برمی‌گردانند
                            try {
                                for (var j=0;j<res.length;j++) {
                                    var e = res[j];
                                    var r = tryExtractReportNoFromEntry(e);
                                    if (r) docs.push(r);
                                }
                                log("getDocumentList: خواندن از متد (object) ", tryNames[i]);
                            } catch(e){}
                        }
                    } catch(e){ log("getDocumentList call error for " + tryNames[i], e); }
                } else if (Array.isArray(fn)) {
                    fn.forEach(function(entry){
                        var r = tryExtractReportNoFromEntry(entry);
                        if (r) docs.push(r);
                    });
                }
            } catch(e){}
        }

        // تلاش 3: بررسی dataSources اگر وجود دارد
        try {
            if (form.dataSources && typeof form.dataSources === "object") {
                Object.keys(form.dataSources).forEach(function(key){
                    try {
                        var ds = form.dataSources[key];
                        if (!ds) return;
                        // اگر دارای Entries یا Items
                        if (Array.isArray(ds)) {
                            ds.forEach(function(entry){ var r = tryExtractReportNoFromEntry(entry); if (r) docs.push(r); });
                        } else {
                            if (ds.Entries && Array.isArray(ds.Entries)) ds.Entries.forEach(function(entry){ var r = tryExtractReportNoFromEntry(entry); if (r) docs.push(r); });
                            if (ds.Items && Array.isArray(ds.Items)) ds.Items.forEach(function(entry){ var r = tryExtractReportNoFromEntry(entry); if (r) docs.push(r); });
                        }
                    } catch(e){}
                });
            }
        } catch(e){ log("getDocumentList dataSources error", e); }

        // اگر هیچ نتیجه‌ای نداریم، تلاش برای خواندن کنترل c_MaxRep/c_MinRep (رشته)
        try {
            var maxRep = getControlValue("c_MaxRep");
            if (maxRep) docs.push(maxRep);
            var minRep = getControlValue("c_MinReportNo");
            if (minRep) docs.push(minRep);
        } catch(e){}

        // نهایی‌سازی: trim, unique و حذف خالی
        var normalized = [];
        docs.forEach(function(x){
            try {
                if (!x) return;
                var s = x.toString().trim();
                if (!s) return;
                if (normalized.indexOf(s) === -1) normalized.push(s);
            } catch(e){}
        });
        log("getDocumentList: تعداد اسناد استخراج‌شده:", normalized.length);
        if (DEBUG) log("getDocumentList sample:", normalized.slice(0,10));
        return normalized;
    }

    // ==================== بخش 4: یافتن بزرگ‌ترین شماره برای basePattern ====================
    function findMaxNumberForBase(basePattern) {
        log("findMaxNumberForBase: شروع برای", basePattern);
        var docs = getDocumentList();
        var maxNumber = null;
        var locationCode = getControlValue("c_LocationCode").trim();
        // الگوهایی که ممکنه reportNo رو پیدا کنن:
        var patterns = [
            /^(.+)-0*([0-9]+)-[A-Za-z0-9]+$/, // BASE-0001-G00
            /^(.+)-0*([0-9]+)$/,              // BASE-0001
            /(.+)-0*([0-9]+)-?.*$/            // fallback
        ];

        for (var i=0;i<docs.length;i++){
            var doc = docs[i];
            if (!doc) continue;
            log("بررسی سند[" + i + "]", doc);

            var matched = false;
            for (var p=0;p<patterns.length && !matched;p++){
                try {
                    var m = doc.match(patterns[p]);
                    if (!m) continue;
                    var docBase = (m[1]||"").toString().trim();
                    var docNum = parseInt(m[2],10);
                    if (isNaN(docNum)) continue;

                    log("پارسه با الگو " + p, { base: docBase, num: docNum });

                    // 1) تطبیق مستقیم
                    if (docBase === basePattern) {
                        if (maxNumber === null || docNum > maxNumber) maxNumber = docNum;
                        matched = true;
                        log("✓ تطبیق مستقیم، شماره فعلی برای max:", maxNumber);
                        break;
                    }

                    // 2) حذف locationCode از docBase و مقایسه
                    if (locationCode && docBase.indexOf(locationCode) !== -1) {
                        var withoutLoc = docBase.split(locationCode).join("").trim();
                        // پاکسازی اضافی
                        withoutLoc = withoutLoc.replace(/--+/g,'-').replace(/^-|-$/g,'');
                        if (withoutLoc === basePattern) {
                            if (maxNumber === null || docNum > maxNumber) maxNumber = docNum;
                            matched = true;
                            log("✓ تطبیق پس از حذف locationCode، شماره فعلی برای max:", maxNumber);
                            break;
                        }
                    }

                    // 3) اگر basePattern خودش ممکنه حاوی location باشد، سعی کن variations بسازی
                    if (locationCode) {
                        var variations = [
                            basePattern + locationCode,
                            locationCode + basePattern,
                            basePattern.replace(/^([^-]+)-/, '$1' + locationCode + '-'),
                            basePattern.replace(/-([^-]+)$/, locationCode + '-$1')
                        ];
                        for (var v=0; v<variations.length; v++) {
                            if (docBase === variations[v]) {
                                if (maxNumber === null || docNum > maxNumber) maxNumber = docNum;
                                matched = true;
                                log("✓ تطبیق با variation " + v + ", شماره فعلی برای max:", maxNumber);
                                break;
                            }
                        }
                        if (matched) break;
                    }

                } catch(e){ log("error parsing doc", e); }
            } // end patterns loop

            if (!matched) log("× هیچ تطبیقی برای این سند نیافت", doc);
        } // end docs loop

        log("findMaxNumberForBase: حداکثر نهایی", maxNumber);
        return maxNumber;
    }

    // ==================== بخش 5: تعیین شماره بعدی ====================
    function getNextNumber(basePattern) {
        log("=== getNextNumber for", basePattern);
        var maxExisting = findMaxNumberForBase(basePattern);
        var defaultFromControl = getDefaultReportFromControl(basePattern);
        var defaultFromList = getDefaultReportFromList(basePattern);

        log("maxExisting", maxExisting, "defaultFromControl", defaultFromControl, "defaultFromList", defaultFromList);

        // انتخاب default ترجیحی
        var defaultNum = (defaultFromControl !== null) ? defaultFromControl : ((defaultFromList !== null) ? defaultFromList : null);

        if (maxExisting !== null && defaultNum !== null) {
            if (maxExisting >= defaultNum) return maxExisting + 1;
            return defaultNum;
        } else if (maxExisting !== null) {
            return maxExisting + 1;
        } else if (defaultNum !== null) {
            return defaultNum;
        } else {
            return 1;
        }
    }

    // ==================== بخش 6: بروزرسانی DefaultReport ====================
    function updateDefaultReportMinMax(basePattern) {
        try {
            log("updateDefaultReportMinMax: شروع بروزرسانی برای", basePattern);
            
            // خواندن مقادیر Min/Max از کنترل‌ها
            var maxReportValue = getControlValue("c_MaxRep") || getControlValue("c_MaxReportNo");
            var minReportValue = getControlValue("c_MinReportNo") || getControlValue("c_MinRep");
            
            log("مقادیر خوانده شده - Max:", maxReportValue, "Min:", minReportValue);
            
            if (!maxReportValue && !minReportValue) {
                log("هیچ مقدار Min/Max یافت نشد");
                return false;
            }
            
            // روش‌های مختلف دسترسی به لیست DefaultReport
            var defaultReports = null;
            var listControl = null;
            
            // تلاش 1: دسترسی مستقیم به لیست
            try {
                if (form.DefaultReport && Array.isArray(form.DefaultReport)) {
                    defaultReports = form.DefaultReport;
                    log("DefaultReport از form.DefaultReport یافت شد");
                }
            } catch(e) { log("خطا در form.DefaultReport", e); }
            
            // تلاش 2: دسترسی از طریق کنترل
            if (!defaultReports) {
                try {
                    listControl = form.GetControl("DefaultReport");
                    if (listControl) {
                        if (typeof listControl.GetItems === "function") {
                            defaultReports = listControl.GetItems();
                        } else if (typeof listControl.Items !== "undefined") {
                            defaultReports = listControl.Items;
                        } else if (Array.isArray(listControl)) {
                            defaultReports = listControl;
                        }
                        log("DefaultReport از کنترل یافت شد");
                    }
                } catch(e) { log("خطا در دسترسی به کنترل DefaultReport", e); }
            }
            
            // تلاش 3: متدهای کمکی
            if (!defaultReports) {
                var methodNames = ["GetDefaultReportList", "GetDefaultReports", "getDefaultReportList"];
                for (var m = 0; m < methodNames.length; m++) {
                    try {
                        if (typeof form[methodNames[m]] === "function") {
                            var res = form[methodNames[m]]();
                            if (Array.isArray(res)) {
                                defaultReports = res;
                                log("DefaultReport از متد " + methodNames[m] + " یافت شد");
                                break;
                            }
                        }
                    } catch(e) { log("خطا در متد " + methodNames[m], e); }
                }
            }
            
            // تلاش 4: دسترسی از طریق dataSources
            if (!defaultReports) {
                try {
                    if (form.dataSources && form.dataSources.DefaultReport) {
                        if (Array.isArray(form.dataSources.DefaultReport)) {
                            defaultReports = form.dataSources.DefaultReport;
                        } else if (form.dataSources.DefaultReport.Items) {
                            defaultReports = form.dataSources.DefaultReport.Items;
                        }
                        log("DefaultReport از dataSources یافت شد");
                    }
                } catch(e) { log("خطا در dataSources", e); }
            }
            
            if (!defaultReports || !Array.isArray(defaultReports) || defaultReports.length === 0) {
                log("هیچ DefaultReport قابل دسترسی یافت نشد");
                return false;
            }
            
            log("تعداد آیتم‌های DefaultReport:", defaultReports.length);
            
            // بروزرسانی رکوردهای مطابق با basePattern
            var updatedCount = 0;
            for (var i = 0; i < defaultReports.length; i++) {
                var item = defaultReports[i];
                if (!item) continue;
                
                try {
                    var itemBasePattern = null;
                    
                    // روش‌های مختلف استخراج BasePattern
                    if (typeof item.get_item === "function") {
                        // SharePoint List Item
                        itemBasePattern = item.get_item("BasePattern") || 
                                        item.get_item("Pattern") || 
                                        item.get_item("Title");
                        log("BasePattern از get_item:", itemBasePattern);
                    } else if (typeof item.get_fieldValue === "function") {
                        // InfoPath
                        itemBasePattern = item.get_fieldValue("BasePattern") || 
                                        item.get_fieldValue("Pattern") || 
                                        item.get_fieldValue("Title");
                        log("BasePattern از get_fieldValue:", itemBasePattern);
                    } else if (item.BasePattern || item.Pattern || item.Title) {
                        // آبجکت ساده
                        itemBasePattern = item.BasePattern || item.Pattern || item.Title;
                        log("BasePattern از فیلد مستقیم:", itemBasePattern);
                    }
                    
                    if (!itemBasePattern) {
                        log("BasePattern برای آیتم " + i + " یافت نشد");
                        continue;
                    }
                    
                    itemBasePattern = itemBasePattern.toString().trim();
                    
                    // بررسی تطبیق با basePattern (با انعطاف بیشتر)
                    var isMatch = false;
                    if (itemBasePattern === basePattern) {
                        isMatch = true;
                    } else {
                        // تطبیق بدون در نظر گیری حروف کوچک/بزرگ
                        if (itemBasePattern.toLowerCase() === basePattern.toLowerCase()) {
                            isMatch = true;
                        }
                        // تطبیق partial
                        else if (itemBasePattern.indexOf(basePattern) !== -1 || basePattern.indexOf(itemBasePattern) !== -1) {
                            isMatch = true;
                        }
                    }
                    
                    if (isMatch) {
                        log("آیتم " + i + " مطابقت دارد:", itemBasePattern);
                        
                        // بروزرسانی MaxReportNo
                        if (maxReportValue) {
                            var maxUpdated = false;
                            if (typeof item.set_item === "function") {
                                item.set_item("MaxReportNo", maxReportValue);
                                maxUpdated = true;
                            } else if (typeof item.set_fieldValue === "function") {
                                item.set_fieldValue("MaxReportNo", maxReportValue);
                                maxUpdated = true;
                            } else if (item.MaxReportNo !== undefined) {
                                item.MaxReportNo = maxReportValue;
                                maxUpdated = true;
                            } else {
                                // ایجاد فیلد جدید
                                item.MaxReportNo = maxReportValue;
                                maxUpdated = true;
                            }
                            log("MaxReportNo بروزرسانی شد:", maxUpdated);
                        }
                        
                        // بروزرسانی MinReportNo
                        if (minReportValue) {
                            var minUpdated = false;
                            if (typeof item.set_item === "function") {
                                item.set_item("MinReportNo", minReportValue);
                                minUpdated = true;
                            } else if (typeof item.set_fieldValue === "function") {
                                item.set_fieldValue("MinReportNo", minReportValue);
                                minUpdated = true;
                            } else if (item.MinReportNo !== undefined) {
                                item.MinReportNo = minReportValue;
                                minUpdated = true;
                            } else {
                                // ایجاد فیلد جدید
                                item.MinReportNo = minReportValue;
                                minUpdated = true;
                            }
                            log("MinReportNo بروزرسانی شد:", minUpdated);
                        }
                        
                        updatedCount++;
                        log("آیتم " + i + " بروزرسانی شد:", {
                            basePattern: itemBasePattern,
                            maxReportNo: maxReportValue,
                            minReportNo: minReportValue
                        });
                    }
                } catch(e) {
                    log("خطا در بروزرسانی آیتم " + i + ":", e);
                }
            }
            
            log("updateDefaultReportMinMax: تعداد آیتم‌های بروزرسانی شده:", updatedCount);
            
            // ذخیره تغییرات - روش‌های مختلف
            if (updatedCount > 0) {
                try {
                    // روش 1: متد مخصوص فرم
                    if (typeof form.UpdateDefaultReportList === "function") {
                        form.UpdateDefaultReportList(defaultReports);
                        log("ذخیره از طریق UpdateDefaultReportList");
                    } 
                    // روش 2: متد عمومی ذخیره
                    else if (typeof form.SaveDefaultReports === "function") {
                        form.SaveDefaultReports();
                        log("ذخیره از طریق SaveDefaultReports");
                    }
                    // روش 3: کنترل لیست
                    else if (listControl && typeof listControl.Update === "function") {
                        listControl.Update();
                        log("ذخیره از طریق کنترل لیست");
                    }
                    // روش 4: ذخیره عمومی فرم
                    else if (typeof form.Save === "function") {
                        form.Save();
                        log("ذخیره از طریق form.Save");
                    }
                    
                    // تلاش برای refresh کنترل
                    try {
                        if (listControl && typeof listControl.Refresh === "function") {
                            listControl.Refresh();
                        }
                    } catch(e) { log("خطا در Refresh", e); }
                    
                } catch(e) {
                    log("خطا در ذخیره DefaultReport:", e);
                }
            }
            
            return updatedCount > 0;
            
        } catch(e) {
            log("updateDefaultReportMinMax: خطای کلی", e);
            return false;
        }
    }

    // ==================== بخش 7: بروزرسانی Min/Max ====================
    function updateMinMax(basePattern, currentNumber) {
        var currentMin = getControlValue("c_MinReportNo");
        var currentMax = getControlValue("c_MaxReportNo") || getControlValue("c_MaxRep");
        var minNum = null, maxNum = null;
        
        try {
            var mm = currentMin.match(/-(\d+)-/); 
            if (mm) minNum = parseInt(mm[1], 10);
            var mx = currentMax.match(/-(\d+)-/); 
            if (mx) maxNum = parseInt(mx[1], 10);
        } catch(e) {}
        
        var updated = false;
        
        if (minNum === null || currentNumber < minNum) {
            var newMinReport = basePattern + "-" + pad4(currentNumber) + "-G00";
            setControlValue("c_MinReportNo", newMinReport);
            log("Min updated", newMinReport);
            updated = true;
        }
        
        if (maxNum === null || currentNumber > maxNum) {
            var newMaxReport = basePattern + "-" + pad4(currentNumber) + "-G00";
            setControlValue("c_MaxReportNo", newMaxReport);
            setControlValue("c_MaxRep", newMaxReport);
            log("Max updated", newMaxReport);
            updated = true;
        }
        
        // بروزرسانی DefaultReport هنگامی که Min/Max تغییر کرده‌اند
        if (updated) {
            updateDefaultReportMinMax(basePattern);
        }
    }

    // ==================== بخش 8: ساخت الگوها ====================
    function createBasePattern() {
        var constantPart = getControlValue("c_ConstantPart");
        var contCode = getControlValue("c_ContCode");
        var psCode = getControlValue("c_PSCode");
        var spCode = getControlValue("c_SPCode");
        var maingrCode = getControlValue("c_MaingrCode");
        var tpCode = getControlValue("c_TPCode");
        var basePattern = constantPart + contCode + "-" + psCode + spCode + "-" + maingrCode + tpCode;
        log("createBasePattern", basePattern);
        return basePattern;
    }

    function createFullPattern() {
        var constantPart = getControlValue("c_ConstantPart");
        var locationCode = getControlValue("c_LocationCode");
        var contCode = getControlValue("c_ContCode");
        var psCode = getControlValue("c_PSCode");
        var spCode = getControlValue("c_SPCode");
        var maingrCode = getControlValue("c_MaingrCode");
        var tpCode = getControlValue("c_TPCode");
        var fullPattern = constantPart + locationCode + contCode + "-" + psCode + spCode + "-" + maingrCode + tpCode;
        log("createFullPattern", fullPattern);
        return fullPattern;
    }

    // ==================== بخش 9: اعتبارسنجی ====================
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
        for (var i=0;i<checks.length;i++){
            if (!getControlValue(checks[i].field)) return checks[i].message;
        }
        return null;
    }

    // ==================== بخش 10: تولید شماره گزارش ====================
    function generateReportNumber(showAlerts) {
        log("generateReportNumber start showAlerts=", showAlerts);
        if (busy) {
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(function(){ generateReportNumber(showAlerts); }, 50);
            return;
        }
        busy = true;
        try {
            if (showAlerts) {
                var ve = validateInputs();
                if (ve) { alert(ve); busy=false; return false; }
            }
            var basePattern = createBasePattern();
            var fullPattern = createFullPattern();
            setControlValue("c_pattern", basePattern);
            if (!showAlerts && validateInputs()) { busy=false; return false; }

            var nextNumber = getNextNumber(basePattern);
            updateMinMax(basePattern, nextNumber);

            var padded = pad4(nextNumber);
            var final = fullPattern + "-" + padded + "-G00";
            setControlValue("c_PartNum", padded);
            setControlValue("c_ReportNo", final);
            setControlValue("c_Rev", "G00");

            try {
                var rc = form.GetControl("c_ReportNo");
                if (rc) { 
                    if (typeof rc.SetEnabled === "function") rc.SetEnabled(false); 
                    else if (rc.disabled !== undefined) rc.disabled = true; 
                }
            } catch(e){}

            // بروزرسانی اضافی DefaultReport هنگام اضافه شدن رکورد جدید به DocumentEntryList
            try {
                // چک کردن اینکه آیا رکورد جدید به DocumentEntryList اضافه شده
                if (form.DocumentEntryList && Array.isArray(form.DocumentEntryList)) {
                    var lastEntry = form.DocumentEntryList[form.DocumentEntryList.length - 1];
                    if (lastEntry) {
                        // اگر رکورد جدید اضافه شده، DefaultReport را بروزرسانی کن
                        updateDefaultReportMinMax(basePattern);
                    }
                }
            } catch(e) {
                log("خطا در بررسی DocumentEntryList", e);
            }

            try { 
                if (typeof form.Save === "function") { 
                    form.Save(); 
                    log("form.Save called"); 
                } 
            } catch(e){ 
                log("save error", e); 
            }

            log("generateReportNumber done:", final);
            busy = false;
            return true;
        } catch(e){ 
            log("generateReportNumber fatal", e); 
            busy=false; 
            return false; 
        }
    }

    // ==================== بخش 11: توابع کمکی و تست ====================
    function manualUpdateDefaultReport() {
        try {
            var basePattern = createBasePattern();
            updateDefaultReportMinMax(basePattern);
            log("بروزرسانی دستی DefaultReport انجام شد");
        } catch(e) {
            log("خطا در بروزرسانی دستی", e);
        }
    }
    
    // تابع تست برای بررسی محتویات DefaultReport
    function debugDefaultReport() {
        try {
            log("=== شروع Debug DefaultReport ===");
            
            // تست دسترسی‌های مختلف
            var methods = [
                { name: "form.DefaultReport", getValue: function() { return form.DefaultReport; } },
                { name: "form.GetControl('DefaultReport')", getValue: function() { return form.GetControl("DefaultReport"); } },
                { name: "form.GetDefaultReportList()", getValue: function() { return form.GetDefaultReportList && form.GetDefaultReportList(); } },
                { name: "form.dataSources.DefaultReport", getValue: function() { return form.dataSources && form.dataSources.DefaultReport; } }
            ];
            
            for (var i = 0; i < methods.length; i++) {
                try {
                    var result = methods[i].getValue();
                    if (result) {
                        log(methods[i].name + " موجود است:", typeof result);
                        if (Array.isArray(result)) {
                            log("  - تعداد آیتم‌ها:", result.length);
                            if (result.length > 0) {
                                log("  - نمونه آیتم اول:", result[0]);
                                // بررسی فیلدهای موجود
                                if (result[0]) {
                                    var fields = Object.keys(result[0]);
                                    log("  - فیلدهای موجود:", fields);
                                }
                            }
                        } else if (result.Items) {
                            log("  - دارای Items با تعداد:", result.Items.length);
                        }
                    } else {
                        log(methods[i].name + " موجود نیست یا null است");
                    }
                } catch(e) {
                    log(methods[i].name + " خطا:", e.message);
                }
            }
            
            // بررسی مقادیر Min/Max فعلی
            var currentMax = getControlValue("c_MaxRep") || getControlValue("c_MaxReportNo");
            var currentMin = getControlValue("c_MinReportNo") || getControlValue("c_MinRep");
            log("مقادیر فعلی - Max:", currentMax, "Min:", currentMin);
            
            // بررسی basePattern فعلی
            var currentBasePattern = createBasePattern();
            log("BasePattern فعلی:", currentBasePattern);
            
            log("=== پایان Debug DefaultReport ===");
            
        } catch(e) {
            log("خطا در debugDefaultReport:", e);
        }
    }

    // ==================== بخش 12: اتصال رویدادها ====================
    function attachEvents() {
        var controlNames = ["c_ConstantPart","c_LocationCode","c_MaingrCode","c_ContCode","c_PSCode","c_SPCode","c_TPCode","c_Subject","c_ReportDate","c_MaxRep","c_DefRep"];
        var handler = function(){ clearTimeout(debounceTimer); debounceTimer = setTimeout(function(){ generateReportNumber(false); }, 100); };
        controlNames.forEach(function(n){ try {
            var c = form.GetControl(n);
            if (c) {
                if (c.SelectionChanged && typeof c.SelectionChanged.connect === "function") c.SelectionChanged.connect(handler);
                else if (typeof c.onchange !== "undefined") { var old=c.onchange; c.onchange=function(e){ try{ if(old) old(e); } catch(ex){} handler(); }; }
            }
        } catch(e){ log("attachEvents error for "+n, e); }});
        try {
            var btn = form.GetControl("c_control1");
            if (btn) {
                if (btn.Click && typeof btn.Click.connect === "function") btn.Click.connect(function(){ generateReportNumber(true); });
                else if (typeof btn.onclick !== "undefined") btn.onclick = function(){ generateReportNumber(true); };
            }
        } catch(e){ log("attach button error", e); }
    }

    // ==================== بخش 13: مقداردهی اولیه ====================
    function initialize() {
        log("initialize");
        try { 
            attachEvents(); 
            setTimeout(function(){ generateReportNumber(false); }, 200); 
        } catch(e){ 
            log("init err", e); 
        }
        
        // اضافه کردن توابع به window برای دسترسی خارجی
        window.generateReportNumber = generateReportNumber;
        window.updateDefaultReportMinMax = updateDefaultReportMinMax;
        window.manualUpdateDefaultReport = manualUpdateDefaultReport;
        window.debugDefaultReport = debugDefaultReport;
    }

    // شروع اجرای اسکریپت
    initialize();

})();