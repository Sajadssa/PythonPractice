import React, { useState } from 'react';
import { FileText, Calendar, DollarSign, Database, Users, CheckCircle, Download } from 'lucide-react';

const PipingProposal = () => {
  const [activeSection, setActiveSection] = useState('overview');

  const sections = {
    overview: {
      title: 'خلاصه اجرایی',
      icon: <FileText className="w-5 h-5" />,
      content: (
        <div className="space-y-4">
          <div className="bg-blue-50 p-4 rounded-lg border-r-4 border-blue-500">
            <h3 className="font-bold text-lg mb-2">عنوان پروژه</h3>
            <p className="text-gray-700">طراحی و پیاده‌سازی سیستم یکپارچه مدیریت پایپینگ پالایشگاهی</p>
          </div>
          
          <div className="grid md:grid-cols-3 gap-4">
            <div className="bg-white p-4 rounded-lg shadow-md border-t-4 border-green-500">
              <p className="text-sm text-gray-600 mb-1">حجم پروژه</p>
              <p className="text-2xl font-bold text-green-600">150,000</p>
              <p className="text-xs text-gray-500">اینچ-قطر</p>
            </div>
            <div className="bg-white p-4 rounded-lg shadow-md border-t-4 border-purple-500">
              <p className="text-sm text-gray-600 mb-1">مدت زمان</p>
              <p className="text-2xl font-bold text-purple-600">5 ماه</p>
              <p className="text-xs text-gray-500">طراحی تا استقرار</p>
            </div>
            <div className="bg-white p-4 rounded-lg shadow-md border-t-4 border-orange-500">
              <p className="text-sm text-gray-600 mb-1">واحدهای تحت پوشش</p>
              <p className="text-2xl font-bold text-orange-600">3 واحد</p>
              <p className="text-xs text-gray-500">دفتر فنی، کنترل کیفیت، تست پکیج</p>
            </div>
          </div>

          <div className="bg-gray-50 p-4 rounded-lg">
            <h4 className="font-semibold mb-3">اهداف پروژه:</h4>
            <ul className="space-y-2">
              <li className="flex items-start gap-2">
                <CheckCircle className="w-5 h-5 text-green-500 mt-0.5 flex-shrink-0" />
                <span>ایجاد یک سیستم یکپارچه جهت مدیریت کامل چرخه عمر پایپینگ</span>
              </li>
              <li className="flex items-start gap-2">
                <CheckCircle className="w-5 h-5 text-green-500 mt-0.5 flex-shrink-0" />
                <span>اتصال و هماهنگی واحدهای دفتر فنی، کنترل کیفیت و تست پکیج</span>
              </li>
              <li className="flex items-start gap-2">
                <CheckCircle className="w-5 h-5 text-green-500 mt-0.5 flex-shrink-0" />
                <span>ردیابی و گزارش‌گیری لحظه‌ای از وضعیت پروژه</span>
              </li>
              <li className="flex items-start gap-2">
                <CheckCircle className="w-5 h-5 text-green-500 mt-0.5 flex-shrink-0" />
                <span>کاهش خطاهای انسانی و افزایش کیفیت داده‌ها</span>
              </li>
            </ul>
          </div>
        </div>
      )
    },
    technical: {
      title: 'معماری فنی',
      icon: <Database className="w-5 h-5" />,
      content: (
        <div className="space-y-4">
          <div className="bg-white p-4 rounded-lg shadow-md">
            <h3 className="font-bold text-lg mb-3 text-blue-600">معماری سیستم</h3>
            <div className="grid md:grid-cols-2 gap-4">
              <div className="border-r-4 border-blue-500 pr-4">
                <h4 className="font-semibold mb-2">لایه رابط کاربری (Frontend)</h4>
                <ul className="text-sm space-y-1 text-gray-700">
                  <li>• Microsoft Access 2016/2019</li>
                  <li>• فرم‌های کاربرپسند فارسی</li>
                  <li>• گزارش‌گیری پیشرفته</li>
                  <li>• داشبورد مدیریتی</li>
                </ul>
              </div>
              <div className="border-r-4 border-green-500 pr-4">
                <h4 className="font-semibold mb-2">لایه پایگاه داده (Backend)</h4>
                <ul className="text-sm space-y-1 text-gray-700">
                  <li>• SQL Server 2017/2019</li>
                  <li>• معماری نرمال‌شده</li>
                  <li>• Stored Procedures</li>
                  <li>• پشتیبان‌گیری خودکار</li>
                </ul>
              </div>
            </div>
          </div>

          <div className="bg-white p-4 rounded-lg shadow-md">
            <h3 className="font-bold text-lg mb-3 text-green-600">ماژول‌های سیستم</h3>
            <div className="space-y-3">
              <div className="bg-blue-50 p-3 rounded">
                <h4 className="font-semibold text-blue-700 mb-1">۱. ماژول دفتر فنی</h4>
                <p className="text-sm text-gray-700">ثبت Isometric، MTO، مشخصات فنی پایپ‌ها، مدیریت نقشه‌ها و ریویژن‌ها</p>
              </div>
              <div className="bg-green-50 p-3 rounded">
                <h4 className="font-semibold text-green-700 mb-1">۲. ماژول کنترل کیفیت</h4>
                <p className="text-sm text-gray-700">ثبت بازرسی‌ها، تست‌های NDT، کنترل جوش‌ها، مدیریت NCR، ردیابی مواد</p>
              </div>
              <div className="bg-purple-50 p-3 rounded">
                <h4 className="font-semibold text-purple-700 mb-1">۳. ماژول تست پکیج</h4>
                <p className="text-sm text-gray-700">Hydro Test، Pneumatic Test، Flushing، Pre-commissioning، گواهی‌نامه‌ها</p>
              </div>
              <div className="bg-orange-50 p-3 rounded">
                <h4 className="font-semibold text-orange-700 mb-1">۴. ماژول گزارش‌گیری و مانیتورینگ</h4>
                <p className="text-sm text-gray-700">داشبورد پیشرفت، گزارشات مدیریتی، چارت‌های تحلیلی، پیگیری تاخیرات</p>
              </div>
            </div>
          </div>

          <div className="bg-yellow-50 p-4 rounded-lg border-r-4 border-yellow-500">
            <h4 className="font-semibold mb-2">ویژگی‌های کلیدی:</h4>
            <div className="grid md:grid-cols-2 gap-2 text-sm">
              <div>✓ چند کاربره با سطوح دسترسی</div>
              <div>✓ ثبت تاریخچه تغییرات (Audit Trail)</div>
              <div>✓ گردش کار خودکار (Workflow)</div>
              <div>✓ اعلان‌های سیستمی</div>
              <div>✓ خروجی Excel و PDF</div>
              <div>✓ جستجوی پیشرفته</div>
            </div>
          </div>
        </div>
      )
    },
    timeline: {
      title: 'برنامه زمان‌بندی',
      icon: <Calendar className="w-5 h-5" />,
      content: (
        <div className="space-y-4">
          <div className="bg-gradient-to-l from-blue-50 to-white p-4 rounded-lg border-r-4 border-blue-500">
            <h3 className="font-bold text-lg mb-2">مدت کل پروژه: 5 ماه (20 هفته)</h3>
          </div>

          <div className="space-y-3">
            <div className="bg-white p-4 rounded-lg shadow-md border-r-4 border-blue-400">
              <div className="flex justify-between items-start mb-2">
                <h4 className="font-bold text-blue-700">فاز ۱: تحلیل و طراحی</h4>
                <span className="text-sm bg-blue-100 px-3 py-1 rounded-full">4 هفته</span>
              </div>
              <ul className="text-sm space-y-1 text-gray-700">
                <li>• هفته 1-2: جمع‌آوری نیازمندی‌ها و تحلیل فرآیندها</li>
                <li>• هفته 2-3: طراحی پایگاه داده و معماری سیستم</li>
                <li>• هفته 3-4: طراحی رابط کاربری و تایید نهایی</li>
              </ul>
            </div>

            <div className="bg-white p-4 rounded-lg shadow-md border-r-4 border-green-400">
              <div className="flex justify-between items-start mb-2">
                <h4 className="font-bold text-green-700">فاز ۲: توسعه و پیاده‌سازی</h4>
                <span className="text-sm bg-green-100 px-3 py-1 rounded-full">8 هفته</span>
              </div>
              <ul className="text-sm space-y-1 text-gray-700">
                <li>• هفته 5-6: ایجاد پایگاه داده SQL Server</li>
                <li>• هفته 7-9: توسعه ماژول دفتر فنی</li>
                <li>• هفته 10-11: توسعه ماژول کنترل کیفیت</li>
                <li>• هفته 12: توسعه ماژول تست پکیج و گزارش‌گیری</li>
              </ul>
            </div>

            <div className="bg-white p-4 rounded-lg shadow-md border-r-4 border-purple-400">
              <div className="flex justify-between items-start mb-2">
                <h4 className="font-bold text-purple-700">فاز ۳: تست و اصلاح</h4>
                <span className="text-sm bg-purple-100 px-3 py-1 rounded-full">3 هفته</span>
              </div>
              <ul className="text-sm space-y-1 text-gray-700">
                <li>• هفته 13-14: تست واحد و یکپارچه‌سازی</li>
                <li>• هفته 15: تست قبولی کاربر (UAT)</li>
              </ul>
            </div>

            <div className="bg-white p-4 rounded-lg shadow-md border-r-4 border-orange-400">
              <div className="flex justify-between items-start mb-2">
                <h4 className="font-bold text-orange-700">فاز ۴: آموزش و استقرار</h4>
                <span className="text-sm bg-orange-100 px-3 py-1 rounded-full">3 هفته</span>
              </div>
              <ul className="text-sm space-y-1 text-gray-700">
                <li>• هفته 16-17: آموزش کاربران (3 واحد)</li>
                <li>• هفته 18: استقرار نهایی و Go-Live</li>
              </ul>
            </div>

            <div className="bg-white p-4 rounded-lg shadow-md border-r-4 border-red-400">
              <div className="flex justify-between items-start mb-2">
                <h4 className="font-bold text-red-700">فاز ۵: پشتیبانی اولیه</h4>
                <span className="text-sm bg-red-100 px-3 py-1 rounded-full">2 هفته</span>
              </div>
              <ul className="text-sm space-y-1 text-gray-700">
                <li>• هفته 19-20: پشتیبانی کامل در محل و رفع مشکلات</li>
              </ul>
            </div>
          </div>

          <div className="bg-gray-50 p-4 rounded-lg">
            <h4 className="font-semibold mb-2">نکات مهم:</h4>
            <ul className="text-sm space-y-1 text-gray-700">
              <li>• جلسات هفتگی پیگیری پیشرفت</li>
              <li>• تحویل مرحله‌ای و دریافت بازخورد</li>
              <li>• امکان تنظیم برنامه بر اساس نیاز</li>
            </ul>
          </div>
        </div>
      )
    },
    cost: {
      title: 'برآورد هزینه',
      icon: <DollarSign className="w-5 h-5" />,
      content: (
        <div className="space-y-4">
          <div className="bg-green-50 p-4 rounded-lg border-r-4 border-green-500">
            <h3 className="font-bold text-xl mb-2">هزینه کل پروژه</h3>
            <p className="text-3xl font-bold text-green-600">3,500,000,000 تومان</p>
            <p className="text-sm text-gray-600 mt-1">(سه میلیارد و پانصد میلیون تومان)</p>
          </div>

          <div className="bg-white rounded-lg shadow-md overflow-hidden">
            <div className="bg-gray-700 text-white p-3">
              <h3 className="font-bold">شرح هزینه‌ها</h3>
            </div>
            <div className="divide-y">
              <div className="p-3 flex justify-between items-center hover:bg-gray-50">
                <div>
                  <p className="font-semibold">۱. تحلیل و طراحی سیستم</p>
                  <p className="text-sm text-gray-600">تحلیل نیازمندی‌ها، طراحی پایگاه داده، UI/UX</p>
                </div>
                <p className="font-bold text-blue-600">500,000,000 تومان</p>
              </div>
              
              <div className="p-3 flex justify-between items-center hover:bg-gray-50">
                <div>
                  <p className="font-semibold">۲. توسعه نرم‌افزار</p>
                  <p className="text-sm text-gray-600">برنامه‌نویسی Access، SQL Server، تست واحد</p>
                </div>
                <p className="font-bold text-blue-600">1,500,000,000 تومان</p>
              </div>
              
              <div className="p-3 flex justify-between items-center hover:bg-gray-50">
                <div>
                  <p className="font-semibold">۳. پیاده‌سازی و استقرار</p>
                  <p className="text-sm text-gray-600">نصب، پیکربندی سرور، مهاجرت داده</p>
                </div>
                <p className="font-bold text-blue-600">400,000,000 تومان</p>
              </div>
              
              <div className="p-3 flex justify-between items-center hover:bg-gray-50">
                <div>
                  <p className="font-semibold">۴. آموزش کاربران</p>
                  <p className="text-sm text-gray-600">دوره‌های آموزشی 3 واحد، مستندات فارسی</p>
                </div>
                <p className="font-bold text-blue-600">200,000,000 تومان</p>
              </div>
              
              <div className="p-3 flex justify-between items-center hover:bg-gray-50">
                <div>
                  <p className="font-semibold">۵. مستندسازی</p>
                  <p className="text-sm text-gray-600">مستندات فنی، راهنمای کاربر، دیاگرام‌ها</p>
                </div>
                <p className="font-bold text-blue-600">150,000,000 تومان</p>
              </div>
              
              <div className="p-3 flex justify-between items-center hover:bg-gray-50">
                <div>
                  <p className="font-semibold">۶. پشتیبانی 12 ماه اول</p>
                  <p className="text-sm text-gray-600">رفع باگ، بروزرسانی، پشتیبانی تلفنی و حضوری</p>
                </div>
                <p className="font-bold text-blue-600">600,000,000 تومان</p>
              </div>
              
              <div className="p-3 flex justify-between items-center hover:bg-gray-50">
                <div>
                  <p className="font-semibold">۷. لایسنس‌ها و ابزارها</p>
                  <p className="text-sm text-gray-600">SQL Server, Access Runtime (در صورت نیاز)</p>
                </div>
                <p className="font-bold text-blue-600">150,000,000 تومان</p>
              </div>
            </div>
            <div className="bg-green-100 p-4 border-t-2 border-green-500">
              <div className="flex justify-between items-center">
                <p className="font-bold text-lg">جمع کل:</p>
                <p className="font-bold text-2xl text-green-700">2,000,000,000 تومان</p>
              </div>
            </div>
          </div>

          <div className="grid md:grid-cols-2 gap-4">
            <div className="bg-blue-50 p-4 rounded-lg border-r-4 border-blue-500">
              <h4 className="font-semibold mb-2">شرایط پرداخت پیشنهادی:</h4>
              <ul className="text-sm space-y-1">
                <li>• 30% پیش‌پرداخت (شروع پروژه)</li>
                <li>• 30% پس از فاز توسعه</li>
                <li>• 30% پس از استقرار نهایی</li>
                <li>• 10% پس از پایان پشتیبانی اولیه</li>
              </ul>
            </div>
            
            <div className="bg-yellow-50 p-4 rounded-lg border-r-4 border-yellow-500">
              <h4 className="font-semibold mb-2">خدمات پس از پشتیبانی اولیه:</h4>
              <ul className="text-sm space-y-1">
                <li>• پشتیبانی سالیانه: 400M تومان</li>
                <li>• تغییرات جزئی: 50M تومان/ماه</li>
                <li>• توسعه ماژول جدید: براساس توافق</li>
              </ul>
            </div>
          </div>

          <div className="bg-gray-100 p-4 rounded-lg text-sm text-gray-700">
            <p className="font-semibold mb-1">⚠️ توجه:</p>
            <p>هزینه‌های فوق بر اساس نرخ‌های جاری بازار ایران (آبان 1404) و با احتساب نوسانات ارزی برآورد شده است. قیمت نهایی پس از بررسی دقیق‌تر نیازمندی‌ها قابل تعدیل است.</p>
          </div>
        </div>
      )
    },
    team: {
      title: 'تیم پروژه',
      icon: <Users className="w-5 h-5" />,
      content: (
        <div className="space-y-4">
          <div className="grid md:grid-cols-2 gap-4">
            <div className="bg-white p-4 rounded-lg shadow-md border-t-4 border-blue-500">
              <h4 className="font-bold text-blue-700 mb-3">مدیر پروژه</h4>
              <ul className="text-sm space-y-2 text-gray-700">
                <li>• هماهنگی کلیه فعالیت‌ها</li>
                <li>• ارتباط با کارفرما</li>
                <li>• کنترل زمان و هزینه</li>
                <li>• مدیریت ریسک</li>
              </ul>
            </div>

            <div className="bg-white p-4 rounded-lg shadow-md border-t-4 border-green-500">
              <h4 className="font-bold text-green-700 mb-3">تحلیلگر سیستم (1 نفر)</h4>
              <ul className="text-sm space-y-2 text-gray-700">
                <li>• تحلیل نیازمندی‌ها</li>
                <li>• طراحی فرآیندها</li>
                <li>• مستندسازی</li>
                <li>• تست قبولی</li>
              </ul>
            </div>

            <div className="bg-white p-4 rounded-lg shadow-md border-t-4 border-purple-500">
              <h4 className="font-bold text-purple-700 mb-3">طراح پایگاه داده (1 نفر)</h4>
              <ul className="text-sm space-y-2 text-gray-700">
                <li>• طراحی ERD و Schema</li>
                <li>• بهینه‌سازی کوئری‌ها</li>
                <li>• Stored Procedures</li>
                <li>• استراتژی Backup</li>
              </ul>
            </div>

            <div className="bg-white p-4 rounded-lg shadow-md border-t-4 border-orange-500">
              <h4 className="font-bold text-orange-700 mb-3">توسعه‌دهنده Access (2 نفر)</h4>
              <ul className="text-sm space-y-2 text-gray-700">
                <li>• برنامه‌نویسی VBA</li>
                <li>• طراحی فرم‌ها و گزارش‌ها</li>
                <li>• اتصال به SQL Server</li>
                <li>• رفع اشکال</li>
              </ul>
            </div>

            <div className="bg-white p-4 rounded-lg shadow-md border-t-4 border-red-500">
              <h4 className="font-bold text-red-700 mb-3">متخصص کنترل کیفیت (1 نفر)</h4>
              <ul className="text-sm space-y-2 text-gray-700">
                <li>• تست سیستم</li>
                <li>• تهیه Test Cases</li>
                <li>• گزارش باگ‌ها</li>
                <li>• تایید کیفیت</li>
              </ul>
            </div>

            <div className="bg-white p-4 rounded-lg shadow-md border-t-4 border-yellow-500">
              <h4 className="font-bold text-yellow-700 mb-3">مربی آموزش (1 نفر)</h4>
              <ul className="text-sm space-y-2 text-gray-700">
                <li>• تهیه محتوای آموزشی</li>
                <li>• برگزاری کلاس‌ها</li>
                <li>• راهنمای کاربری</li>
                <li>• پشتیبانی کاربران</li>
              </ul>
            </div>
          </div>

          <div className="bg-blue-50 p-4 rounded-lg border-r-4 border-blue-500">
            <h4 className="font-semibold mb-2">مشاوران تخصصی (در صورت نیاز):</h4>
            <ul className="text-sm space-y-1 text-gray-700">
              <li>• مشاور صنعت پالایشگاهی</li>
              <li>• متخصص امنیت اطلاعات</li>
              <li>• کارشناس Performance Tuning</li>
            </ul>
          </div>
        </div>
      )
    }
  };

  const handleDownload = () => {
    const content = `
===========================================
پروپوزال سیستم مدیریت پایپینگ پالایشگاه
===========================================

تاریخ: ${new Date().toLocaleDateString('fa-IR')}

${Object.entries(sections).map(([key, section]) => `
${section.title}
${'='.repeat(section.title.length)}
[محتوای این بخش در نسخه وب قابل مشاهده است]
`).join('\n')}

-------------------------------------------
تهیه شده توسط: سیستم مدیریت پروژه
-------------------------------------------
    `.trim();

    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Piping-Management-Proposal.txt';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="max-w-6xl mx-auto p-6 bg-gray-50" dir="rtl">
      {/* Header */}
      <div className="bg-gradient-to-l from-blue-600 to-blue-800 text-white p-8 rounded-lg shadow-lg mb-6">
        <h1 className="text-3xl font-bold mb-2">پروپوزال پروژه</h1>
        <h2 className="text-xl mb-4">سیستم یکپارچه مدیریت پایپینگ پالایشگاهی</h2>
        <p className="text-blue-100">Access + SQL Server | حجم: 150,000 اینچ-قطر</p>
      </div>

      {/* Navigation */}
      <div className="bg-white rounded-lg shadow-md mb-6 overflow-x-auto">
        <div className="flex gap-2 p-4">
          {Object.entries(sections).map(([key, section]) => (
            <button
              key={key}
              onClick={() => setActiveSection(key)}
              className={`flex items-center gap-2 px-4 py-2 rounded-lg transition-all whitespace-nowrap ${
                activeSection === key
                  ? 'bg-blue-600 text-white shadow-md'
                  : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
              }`}
            >
              {section.icon}
              <span>{section.title}</span>
            </button>
          ))}
        </div>
      </div>

      {/* Content */}
      <div className="bg-white rounded-lg shadow-md p-6 mb-6">
        {sections[activeSection].content}
      </div>

      {/* Footer */}
      <div className="bg-white rounded-lg shadow-md p-6">
        <div className="flex flex-col md:flex-row justify-between items-center gap-4">
          <div className="text-center md:text-right">
            <p className="text-gray-600 text-sm mb-1">برای دریافت نسخه کامل PDF یا مشاوره رایگان:</p>
            <p className="font-semibold text-blue-600">تماس: 09917540483 | ایمیل: sajadsaeediazad0007@gmail.com</p>
          </div>
          <button
            onClick={handleDownload}
            className="flex items-center gap-2 bg-green-600 hover:bg-green-700 text-white px-6 py-3 rounded-lg shadow-md transition-all"
          >
            <Download className="w-5 h-5" />
            <span>دانلود خلاصه پروپوزال</span>
          </button>
        </div>
        
        <div className="mt-6 pt-6 border-t border-gray-200 text-center text-sm text-gray-500">
          <p>این پروپوزال محرمانه بوده و تنها برای استفاده سازمان درخواست‌کننده می‌باشد</p>
          <p className="mt-1">اعتبار پروپوزال: 60 روز از تاریخ صدور</p>
        </div>
      </div>
    </div>
  );
};

export default PipingProposal;