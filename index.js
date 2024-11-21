const { Client, LocalAuth } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const ExcelJS = require('exceljs');
require('dotenv').config();

// إعداد العميل
const client = new Client({
    authStrategy: new LocalAuth(),
});

// تحميل البيانات من ملف Excel
let data = [];
async function loadDataFromExcel() {
    try {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('gas18-11-2024.xlsx');
        const worksheet = workbook.worksheets[0];

        worksheet.eachRow((row) => {
            const idNumber = row.getCell(1).value?.toString().trim();
            const name = row.getCell(2).value?.toString().trim();
            const province = row.getCell(3).value?.toString().trim();
            const district = row.getCell(4).value?.toString().trim();
            const area = row.getCell(5).value?.toString().trim();
            const distributorId = row.getCell(6).value?.toString().trim();
            const distributorName = row.getCell(7).value?.toString().trim();
            const distributorPhone = row.getCell(8).value?.toString().trim();
            const status = row.getCell(9).value?.toString().trim();
            const orderDate = row.getCell(12).value?.toString().trim();

            if (idNumber && name) {
                data.push({
                    idNumber,
                    name,
                    province: province || "غير متوفر",
                    district: district || "غير متوفر",
                    area: area || "غير متوفر",
                    distributorId: distributorId || "غير متوفر",
                    distributorName: distributorName || "غير متوفر",
                    distributorPhone: distributorPhone || "غير متوفر",
                    status: status || "غير متوفر",
                    orderDate: orderDate || "غير متوفر",
                });
            }
        });
        console.log('تم تحميل البيانات بنجاح.');
    } catch (error) {
        console.error('حدث خطأ أثناء قراءة ملف Excel:', error.message);
    }
}

// قراءة البيانات عند بدء البوت
loadDataFromExcel();

// التعامل مع الرسائل الواردة
client.on('message', async (message) => {
    const input = message.body.trim();
    const chatId = message.from;

    if (input === "🔍 البحث برقم الهوية أو الاسم") {
        client.sendMessage(chatId, "📝 أدخل رقم الهوية أو الاسم للبحث:");
    } else {
        const user = data.find(entry => entry.idNumber === input || entry.name === input);
        if (user) {
            const response = `
🔍 **تفاصيل الطلب:**

👤 **الاسم**: ${user.name}
🏘️ **الحي / المنطقة**: ${user.area}
🏙️ **المدينة**: ${user.district}
📍 **المحافظة**: ${user.province}

📛 **اسم الموزع**: ${user.distributorName}
📞 **رقم جوال الموزع**: ${user.distributorPhone}
🆔 **هوية الموزع**: ${user.distributorId}

📜 **الحالة**: ${user.status}
📅 **تاريخ الطلب**: ${user.orderDate}
            `;
            client.sendMessage(chatId, response);
        } else {
            client.sendMessage(chatId, "⚠️ لم أتمكن من العثور على بيانات للمدخل المقدم.");
        }
    }
});

// بدء البوت
client.on('qr', (qr) => {
    qrcode.generate(qr, { small: true });
});
client.on('ready', () => {
    console.log('Bot is ready!');
});
client.initialize();
