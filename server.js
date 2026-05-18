// server.js 完整代码（新增考勤记录全功能）
const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const app = express();
const port = 3000;

app.use(bodyParser.json());

// -------------------------- 配置文件路径 --------------------------
// 员工信息Excel（你原来的HR.xlsx）
const HR_FILE_PATH = path.join(__dirname, 'HR.xlsx');
// 考勤记录Excel（自动创建，用来存所有打卡记录）
const ATTENDANCE_FILE_PATH = path.join(__dirname, 'attendance.xlsx');

// -------------------------- 工具函数：初始化考勤Excel --------------------------
function initAttendanceExcel() {
  // 如果考勤文件不存在，自动创建
  if (!fs.existsSync(ATTENDANCE_FILE_PATH)) {
    const workbook = xlsx.utils.book_new();
    // 创建考勤记录表头
    const header = [
      '打卡ID', '用户代码', '用户姓名', '所属分部',
      '打卡时间', '打卡纬度', '打卡经度', '打卡照片数量'
    ];
    const sheet = xlsx.utils.aoa_to_sheet([header]);
    xlsx.utils.book_append_sheet(workbook, sheet, '考勤记录');
    xlsx.writeFile(workbook, ATTENDANCE_FILE_PATH);
    console.log('✅ 考勤记录Excel文件已自动创建');
  }
}

// -------------------------- 工具函数：读取员工信息 --------------------------
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

// -------------------------- 工具函数：新增考勤记录 --------------------------
function addAttendanceRecord(record) {
  try {
    // 读取现有考勤文件
    const workbook = xlsx.readFile(ATTENDANCE_FILE_PATH);
    const sheet = workbook.Sheets['考勤记录'];
    const data = xlsx.utils.sheet_to_json(sheet);
    
    // 生成唯一打卡ID（时间戳+随机数）
    const recordId = Date.now() + '-' + Math.floor(Math.random() * 1000);
    // 新增记录
    const newRecord = {
      '打卡ID': recordId,
      '用户代码': record.userCode,
      '用户姓名': record.userName,
      '所属分部': record.userDept,
      '打卡时间': new Date().toLocaleString('zh-CN'), // 本地时间格式
      '打卡纬度': record.latitude,
      '打卡经度': record.longitude,
      '打卡照片数量': record.photoCount
    };
    
    data.push(newRecord);
    // 写回Excel
    const newSheet = xlsx.utils.json_to_sheet(data);
    workbook.Sheets['考勤记录'] = newSheet;
    xlsx.writeFile(workbook, ATTENDANCE_FILE_PATH);
    
    console.log('✅ 考勤记录已保存，打卡ID：', recordId);
    return { success: true, recordId };
  } catch (error) {
    console.error('❌ 考勤记录保存失败：', error);
    return { success: false, error: error.message };
  }
}

// -------------------------- 工具函数：获取用户历史考勤记录 --------------------------
function getUserAttendanceHistory(userCode) {
  try {
    const workbook = xlsx.readFile(ATTENDANCE_FILE_PATH);
    const sheet = workbook.Sheets['考勤记录'];
    const data = xlsx.utils.sheet_to_json(sheet);
    
    // 筛选该用户的所有记录，按时间倒序
    const userRecords = data
      .filter(item => item['用户代码'] === userCode)
      .sort((a, b) => new Date(b['打卡时间']) - new Date(a['打卡时间']));
    
    console.log(`✅ 已获取用户${userCode}的历史考勤记录，共${userRecords.length}条`);
    return { success: true, records: userRecords };
  } catch (error) {
    console.error('❌ 历史考勤记录读取失败：', error);
    return { success: false, error: error.message };
  }
}

// -------------------------- 接口1：登录校验 --------------------------
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

// -------------------------- 接口2：提交考勤打卡 --------------------------
app.post('/api/attendance/submit', (req, res) => {
  const { userCode, userName, userDept, latitude, longitude, photoCount } = req.body;
  
  console.log('📥 收到打卡提交请求：', { userCode, userName, latitude, longitude });
  
  // 非空校验
  if (!userCode || !userName || !latitude || !longitude) {
    return res.json({ code: 400, msg: '打卡数据不完整，请检查' });
  }

  // 新增考勤记录
  const result = addAttendanceRecord({
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

// -------------------------- 接口3：获取用户历史考勤记录 --------------------------
app.get('/api/attendance/history', (req, res) => {
  const { userCode } = req.query;
  
  console.log('📥 收到历史记录查询请求：用户代码=', userCode);
  
  if (!userCode) {
    return res.json({ code: 400, msg: '用户代码不能为空' });
  }

  const result = getUserAttendanceHistory(userCode);

  if (result.success) {
    res.json({ code: 200, msg: '获取成功', data: { records: result.records } });
  } else {
    res.json({ code: 500, msg: '获取历史记录失败' });
  }
});

// -------------------------- 启动服务 --------------------------
// 初始化考勤Excel
initAttendanceExcel();
// 启动服务
app.listen(port, () => {
  console.log(`✅ 考勤打卡后端服务已启动：http://localhost:${port}`);
  console.log(`📌 登录接口：http://localhost:${port}/api/login`);
  console.log(`📌 打卡提交接口：http://localhost:${port}/api/attendance/submit`);
  console.log(`📌 历史记录接口：http://localhost:${port}/api/attendance/history`);
  // 预读取员工信息
  getAccountMap();
});