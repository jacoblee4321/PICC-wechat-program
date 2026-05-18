// 完整适配Upstash Redis的考勤后端代码，直接复制替换即可
const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const path = require('path');
const { Redis } = require('@upstash/redis');
const app = express();

// 1. 核心配置：适配Vercel自动分配的端口
const PORT = process.env.PORT || 3000;

// 2. 初始化Upstash Redis连接（自动读取Vercel环境变量，不用手动改）
const redis = new Redis({
  url: process.env.KV_REST_API_URL,
  token: process.env.KV_REST_API_TOKEN,
});

// 3. 中间件配置：解析JSON+微信小程序跨域适配
app.use(bodyParser.json());
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// 4. 只读文件配置：HR.xlsx是部署时上传的，只读文件系统可正常读取
const HR_FILE_PATH = path.join(__dirname, 'HR.xlsx');

// 5. 工具函数：读取员工信息（和你原来的逻辑完全一致）
function getAccountMap() {
  try {
    const workbook = xlsx.readFile(HR_FILE_PATH);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = xlsx.utils.sheet_to_json(sheet);
    
    const accountMap = {};
    data.forEach(item => {
      const nameKey = Object.keys(item).find(key => key.includes('姓名'));
      const codeKey = Object.keys(item).find(key => key.includes('代码'));
      const deptKey = Object.keys(item).find(key => key.includes('分部'));
      
      if (!nameKey || !codeKey || !deptKey) return;
      
      const name = String(item[nameKey]).trim();
      const code = String(item[codeKey]).trim().replace(/\s/g, '');
      const dept = String(item[deptKey]).trim();
      
      if (!code || !name || !dept) return;
      
      accountMap[code] = { name, code, dept };
    });

    console.log('✅ 共读取到有效员工账号：', Object.keys(accountMap).length, '条');
    return accountMap;
  } catch (error) {
    console.error('❌ 员工信息Excel读取失败：', error);
    return {};
  }
}

// 6. 工具函数：新增考勤记录到Upstash Redis
async function addAttendanceRecord(record) {
  try {
    // 生成唯一打卡ID
    const recordId = Date.now() + '-' + Math.floor(Math.random() * 1000);
    // 构造考勤记录
    const newRecord = {
      recordId,
      userCode: record.userCode,
      userName: record.userName,
      userDept: record.userDept,
      checkInTime: new Date().toLocaleString('zh-CN'),
      latitude: record.latitude,
      longitude: record.longitude,
      photoCount: record.photoCount || 0
    };

    // 1. 存单条打卡记录
    await redis.set(`attendance:${recordId}`, JSON.stringify(newRecord));
    // 2. 把记录ID添加到该用户的打卡列表，方便后续查询
    const userRecordsKey = `user:${record.userCode}:attendance`;
    const userRecords = JSON.parse(await redis.get(userRecordsKey) || '[]');
    userRecords.unshift(recordId); // 最新记录放最前面
    await redis.set(userRecordsKey, JSON.stringify(userRecords));

    console.log('✅ 考勤记录已保存，打卡ID：', recordId);
    return { success: true, recordId };
  } catch (error) {
    console.error('❌ 考勤记录保存失败：', error);
    return { success: false, error: error.message };
  }
}

// 7. 工具函数：获取用户历史考勤记录
async function getUserAttendanceHistory(userCode) {
  try {
    // 1. 先获取该用户的所有打卡记录ID
    const userRecordsKey = `user:${userCode}:attendance`;
    const recordIds = JSON.parse(await redis.get(userRecordsKey) || '[]');
    
    // 2. 批量获取所有打卡记录详情
    const records = [];
    for (const recordId of recordIds) {
      const recordStr = await redis.get(`attendance:${recordId}`);
      if (recordStr) records.push(JSON.parse(recordStr));
    }

    console.log(`✅ 已获取用户${userCode}的历史考勤记录，共${records.length}条`);
    return { success: true, records };
  } catch (error) {
    console.error('❌ 历史考勤记录读取失败：', error);
    return { success: false, error: error.message };
  }
}

// 8. 接口1：登录校验（和你原来的逻辑完全一致，小程序不用改）
app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  const inputCode = String(username).trim().replace(/\s/g, '');
  const inputPwd = String(password).trim().replace(/\s/g, '');
  
  console.log('📥 收到登录请求：代码=', inputCode);

  const accountMap = getAccountMap();
  const userInfo = accountMap[inputCode];

  if (!userInfo) {
    console.log('❌ 登录失败：账号不存在');
    return res.json({ code: 400, msg: '该账号不存在，请检查输入的代码' });
  }
  if (userInfo.code !== inputPwd) {
    console.log('❌ 登录失败：密码错误');
    return res.json({ code: 400, msg: '密码错误，请输入和账号一致的代码' });
  }

  console.log('✅ 登录成功！姓名：', userInfo.name, '代码：', userInfo.code, '分部：', userInfo.dept);
  res.json({ 
    code: 200, 
    msg: '登录成功',
    data: {
      userName: userInfo.name,
      userCode: userInfo.code,
      userDept: userInfo.dept
    }
  });
});

// 9. 接口2：提交考勤打卡（适配Redis，小程序不用改）
app.post('/api/attendance/submit', async (req, res) => {
  const { userCode, userName, userDept, latitude, longitude, photoCount } = req.body;
  
  console.log('📥 收到打卡提交请求：', { userCode, userName, latitude, longitude });
  
  // 非空校验
  if (!userCode || !userName || !latitude || !longitude) {
    return res.json({ code: 400, msg: '打卡数据不完整，请检查' });
  }

  // 新增考勤记录
  const result = await addAttendanceRecord({
    userCode,
    userName,
    userDept,
    latitude,
    longitude,
    photoCount: photoCount || 0
  });

  if (result.success) {
    res.json({ code: 200, msg: '打卡成功', data: { recordId: result.recordId } });
  } else {
    res.json({ code: 500, msg: '打卡失败，服务器保存记录出错' });
  }
});

// 10. 接口3：获取用户历史考勤记录（适配Redis，小程序不用改）
app.get('/api/attendance/history', async (req, res) => {
  const { userCode } = req.query;
  
  console.log('📥 收到历史记录查询请求：用户代码=', userCode);
  
  if (!userCode) {
    return res.json({ code: 400, msg: '用户代码不能为空' });
  }

  const result = await getUserAttendanceHistory(userCode);

  if (result.success) {
    res.json({ code: 200, msg: '获取成功', data: { records: result.records } });
  } else {
    res.json({ code: 500, msg: '获取历史记录失败' });
  }
});

// 11. 健康检查接口：验证服务是否正常运行
app.get('/', (req, res) => {
  res.send('✅ 考勤打卡后端服务正常运行！已适配Upstash Redis');
});

// 12. 初始化服务：预读取员工信息
getAccountMap();

// 13. 启动服务：适配Vercel
app.listen(PORT, () => {
  console.log(`✅ 考勤打卡服务已启动，端口：${PORT}`);
});

// 14. Vercel Serverless函数必须导出app
module.exports = app;
