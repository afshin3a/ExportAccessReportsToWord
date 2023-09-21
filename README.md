**Export Report (MS-Access) to Word - VBA Macro**
**English**

This VBA code is designed to create an Access report in DOCX (Word) format using a button on an Access form. The Access report is converted to an RTF (Rich Text Format) file and then saved as a DOCX format using Microsoft Word.

**How to Use**

Setting up the code in Access:

Add the above code to a module in your Access database.
Create an Access form and add a button named btnExportToWord to it.
Go back to edit the code and specify the name of your desired report in the reportName variable.

**Using the code:**

Open your Access database and open a form that contains the btnExportToWord button.
Click the button to execute the code.
A file save location window will open. Choose the location and name for the output DOCX file.
After processing is complete, a successful message will be displayed.
Note
This code assumes that Microsoft Word is installed. Otherwise, you may encounter errors.

The code saves the temporary RTF file in the Temp directory on your computer and then deletes the temporary RTF file after creating the DOCX file.

You can use this code to create Word-format reports from other reports in Access by simply changing the name of the report in the reportName variable.

**Credits**

This code was developed by Asad Afshin, and you can find newer versions or possible improvements by visiting  https://github.com/afshin3a/ExportAccessReportsToWord 

**تبدیل ریپورت به ورد در اکسس- ماکروی VBA**

**Persian (فارسی)**

این کد VBA برای ایجاد یک گزارش Access به فرمت DOCX (Word) با استفاده از یک دکمه در فرم Access طراحی شده است. گزارش Access به یک فایل RTF (Rich Text Format) تبدیل شده و سپس با استفاده از Microsoft Word به فرمت DOCX ذخیره می‌شود.

**نحوه استفاده**

راه اندازی کد در Access:

کد بالا را به یک ماژول در پایگاه داده Access خود اضافه کنید.
یک فرم Access ایجاد کنید و یک دکمه به نام btnExportToWord به آن اضافه کنید.
دوباره به ویرایش کد بروید و نام گزارش مورد نظر خود را در متغیر reportName تعیین کنید.

**استفاده از کد:**

پایگاه داده Access را باز کرده و فرمی را باز کنید که دکمه btnExportToWord در آن وجود دارد.
بر روی دکمه کلیک کنید تا کد اجرا شود.
یک پنجره انتخاب مکان ذخیره فایل باز می‌شود. مکان و نام فایل خروجی DOCX را انتخاب کنید.
پس از اتمام پردازش، یک پیغام موفقیت آمیز نمایش داده می‌شود.
**نکته**

- این کد فرض می‌کند که Microsoft Word نصب شده است. در غیر این صورت، ممکن است با خطا مواجه شوید.

- این کد فایل RTF موقت را در مسیر Temp کامپیوتر شما ذخیره می‌کند و سپس پس از ایجاد فایل DOCX، فایل RTF موقت را حذف می‌کند.

- می‌توانید از این کد برای ایجاد گزارش‌های به فرمت Word از دیگر گزارش‌های Access استفاده کنید، فقط کافیست نام گزارش مورد نظر خود را در متغیر reportName تغییر دهید.

**اعتبار**

این کد توسط اسد افشین توسعه داده شده و می‌توانید با مراجعه به https://github.com/afshin3a/ExportAccessReportsToWord نسخه‌های جدیدتر یا بهبود‌های ممکن را پیدا کنید.
