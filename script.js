/**
 * ASEAN Olympic & Paralympic Dashboard
 * ข้อมูลสถิตินักกีฬาผู้ได้รับเหรียญรางวัลในภูมิภาคเอเชียตะวันออกเฉียงใต้
 * ข้อมูลจริงจากไฟล์ olympic.xlsx (ทั้งหมด 334 รายการ)
 */

// ==========================================================================
// 1. ชุดข้อมูลจริง (334 รายการ)
// ==========================================================================
const OLYMPIC_DATA = [
  {
    "number": 1,
    "year": 1976,
    "city": "montreal",
    "name": "พเยาว์ พูลธรัตน์",
    "age": "18",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นไลต์ฟลายเวต (ไม่เกิน 48 กก.)",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 2,
    "year": 1984,
    "city": "los angeles",
    "name": "ทวี อัมพรมหา",
    "age": "24",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นไลต์เวลเตอร์เวต",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 3,
    "year": 1988,
    "city": "Seoul",
    "name": "ผจญ มูลสัน",
    "age": "20",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นแบนตัมเวต",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 4,
    "year": 1992,
    "city": "Barcelona",
    "name": "อาคม เฉ่งไล่",
    "age": "22",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นเวลเตอร์เวต",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 5,
    "year": 1996,
    "city": "Atlanta",
    "name": "วิชัย ราชานนท์",
    "age": "27",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นแบนตัมเวต",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 6,
    "year": 1996,
    "city": "Atlanta",
    "name": "สมรักษ์ คำสิงห์",
    "age": "23",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นเฟเธอร์เวต",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 7,
    "year": 2000,
    "city": "Sydney",
    "name": "เกษราภรณ์ สุตา",
    "age": "28",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักหญิง รุ่น 58 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 8,
    "year": 2000,
    "city": "Sydney",
    "name": "พรชัย ทองบุราณ",
    "age": "26",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นไลต์มิดเดิลเวต",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 9,
    "year": 2000,
    "city": "Sydney",
    "name": "วิจารณ์ พลฤทธิ์",
    "age": "24",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นฟลายเวต",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 10,
    "year": 2004,
    "city": "Athens",
    "name": "มนัส บุญจำนงค์",
    "age": "24",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นไลต์เวลเตอร์เวต",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 11,
    "year": 2004,
    "city": "Athens",
    "name": "เยาวภา บุรพลชัย",
    "age": "19",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "เทควันโด",
    "sportSub": "เทควันโดหญิง รุ่น 49 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 12,
    "year": 2004,
    "city": "Athens",
    "name": "สุริยา ปราสาทหินพิมาย",
    "age": "24",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นมิดเดิลเวต",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 13,
    "year": 2004,
    "city": "Athens",
    "name": "วันดี คำเอี่ยม",
    "age": "26",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักหญิง รุ่น 58 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 14,
    "year": 2004,
    "city": "Athens",
    "name": "อารีย์รัตน วิรัฐถาวร",
    "age": "24",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักหญิง รุ่น 48 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 15,
    "year": 2004,
    "city": "Athens",
    "name": "วรพจน์ เพชรขุ้ม",
    "age": "23",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นแบนตัมเวต",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 16,
    "year": 2004,
    "city": "Athens",
    "name": "ปวีณา ทองสุก",
    "age": "25",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักหญิง รุ่น 75 กก.",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 17,
    "year": 2004,
    "city": "Athens",
    "name": "อุดมพร พลศักดิ์",
    "age": "22",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักหญิง รุ่น 53 กก.",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 18,
    "year": 2008,
    "city": "Beijing",
    "name": "สมจิตร จงจอหอ",
    "age": "33",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นฟลายเวต",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 19,
    "year": 2008,
    "city": "Beijing",
    "name": "ประภาวดี เจริญรัตนธารากูล",
    "age": "25",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักหญิง รุ่น 53 กก.",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 20,
    "year": 2008,
    "city": "Beijing",
    "name": "มนัส บุญจำนงค์",
    "age": "28",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นไลต์เวลเตอร์เวต",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 21,
    "year": 2008,
    "city": "Beijing",
    "name": "บุตรี เผือดผ่อง",
    "age": "17",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "เทควันโด",
    "sportSub": "เทควันโดหญิง รุ่น 49 กก.",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 22,
    "year": 2008,
    "city": "Beijing",
    "name": "เพ็ญศิริ เหล่าศิริกุล",
    "age": "24",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักหญิง รุ่น 48 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 23,
    "year": 2008,
    "city": "Beijing",
    "name": "วันดี คำเอี่ยม",
    "age": "30",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักหญิง รุ่น 58 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 24,
    "year": 2012,
    "city": "London",
    "name": "พิมศิริ ศิริแก้ว",
    "age": "22",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักหญิง รุ่น 58 กก.",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 25,
    "year": 2012,
    "city": "London",
    "name": "แก้ว พงษ์ประยูร",
    "age": "31",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นไลต์ฟลายเวต",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 26,
    "year": 2012,
    "city": "London",
    "name": "ชนาธิป ซ้อนขำ",
    "age": "21",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "เทควันโด",
    "sportSub": "เทควันโดหญิง รุ่น 49 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 27,
    "year": 2012,
    "city": "London",
    "name": "ศิริภุช กุลน้อย",
    "age": "19",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักหญิง รุ่น 58 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 28,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "โสภิตา ธนสาร",
    "age": "21",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักหญิง รุ่น 48 กก.",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 29,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "สุกัญญา ศรีสุราช",
    "age": "21",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักหญิง รุ่น 58 กก.",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 30,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "พิมศิริ ศิริแก้ว",
    "age": "26",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักหญิง รุ่น 58 กก.",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 31,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "เทวินทร์ หาญปราบ",
    "age": "18",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "เทควันโด",
    "sportSub": "เทควันโดชาย รุ่น 58 กก.",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 32,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "พาณิภัค วงศ์พัฒนกิจ",
    "age": "19",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "เทควันโด",
    "sportSub": "เทควันโดหญิง รุ่น 49 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 33,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "สินธุ์เพชร กรวยทอง",
    "age": "20",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักชาย รุ่น 56 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 34,
    "year": 2020,
    "city": "Tokyo",
    "name": "พาณิภัค วงศ์พัฒนกิจ",
    "age": "23",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "เทควันโด",
    "sportSub": "เทควันโดหญิง รุ่น 49 กก.",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 35,
    "year": 2020,
    "city": "Tokyo",
    "name": "สุดาพร สีสอนดี",
    "age": "29",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นไลต์เวต",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 36,
    "year": 2024,
    "city": "Paris",
    "name": "พาณิภัค วงศ์พัฒนกิจ",
    "age": "26",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "เทควันโด",
    "sportSub": "เทควันโดหญิง รุ่น 49 กก.",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 37,
    "year": 2024,
    "city": "Paris",
    "name": "กุลวุฒิ วิทิตศานต์",
    "age": "23",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 38,
    "year": 2024,
    "city": "Paris",
    "name": "ธีรพงศ์ ศิลาชัย",
    "age": "20",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักชาย รุ่น 61 กก.",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 39,
    "year": 2024,
    "city": "Paris",
    "name": "วีรพล วิชุมา",
    "age": "19",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักชาย รุ่น 73 กก.",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 40,
    "year": 2024,
    "city": "Paris",
    "name": "สุรจนา คำเบ้า",
    "age": "25",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ยกน้ำหนักหญิง รุ่น 49 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 41,
    "year": 2024,
    "city": "Paris",
    "name": "จันทร์แจ่ม สุวรรณเพ็ง",
    "age": "23",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นเวลเตอร์เวต",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 42,
    "year": 1988,
    "city": "Seoul",
    "name": "สกุล คำตัน",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "พุ่งแหลนชาย",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 43,
    "year": 1992,
    "city": "Barcelona",
    "name": "บูชิต อังคุระนาวิน",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "พุ่งแหลนชาย",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 44,
    "year": 1996,
    "city": "Atlanta",
    "name": "ประสบโชค กลั่นเงิน",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 10,000 เมตร",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 45,
    "year": 1996,
    "city": "Atlanta",
    "name": "พนม ลักษณะพริ้ม",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "กบ 100 เมตร",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 46,
    "year": 2000,
    "city": "Sydney",
    "name": "ประวัติ วะโฮรัมย์",
    "age": "19",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 10,000 เมตร T54",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 47,
    "year": 2000,
    "city": "Sydney",
    "name": "ประวัติ วะโฮรัมย์",
    "age": "19",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 5,000 เมตร T54",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 48,
    "year": 2000,
    "city": "Sydney",
    "name": "ศุภชัย โกยทรัพย์",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 200 เมตร T54",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 49,
    "year": 2000,
    "city": "Sydney",
    "name": "ศุภชัย โกยทรัพย์, โสภา อินทเสน, ประสิทธิ์ ทองชื่น, อำไพ เสือเหลือง",
    "age": "252",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 4×100 เมตร T54",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 50,
    "year": 2000,
    "city": "Sydney",
    "name": "สมชาย ดวงแก้ว",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ผีเสื้อ 50 เมตร S4",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 51,
    "year": 2000,
    "city": "Sydney",
    "name": "ศุภชัย โกยทรัพย์, โสภา อินทเสน, ประสิทธิ์ ทองชื่น, อำไพ เสือเหลือง",
    "age": "252",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 4×400 เมตร T54",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 52,
    "year": 2000,
    "city": "Sydney",
    "name": "โสภา อินทเสน",
    "age": "25",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 100 เมตร T53",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 53,
    "year": 2000,
    "city": "Sydney",
    "name": "สมชาย ดวงแก้ว",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "เดี่ยวผสม 150 เมตร SM3",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 54,
    "year": 2000,
    "city": "Sydney",
    "name": "สมควร อนนท์",
    "age": "25",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "รุ่นไม่เกิน 56 กก.",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 55,
    "year": 2000,
    "city": "Sydney",
    "name": "ทองสา มารศรี",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "รุ่นไม่เกิน 48 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 56,
    "year": 2000,
    "city": "Sydney",
    "name": "สายฝน แก้วศรี",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "กบ 50 เมตร SB2",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 57,
    "year": 2004,
    "city": "Athens",
    "name": "ประวัติ วะโฮรัมย์, พิเชษฐ์ กรุงเกตุ, ศุภชัย โกยทรัพย์, เรวัติ ต๋านะ",
    "age": "252",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 4×100 เมตร",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 58,
    "year": 2004,
    "city": "Athens",
    "name": "ประวัติ วะโฮรัมย์, พิเชษฐ์ กรุงเกตุ, ศุภชัย โกยทรัพย์, เรวัติ ต๋านะ",
    "age": "252",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 4×400 เมตร",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 59,
    "year": 2004,
    "city": "Athens",
    "name": "สายสุนีย์ จ๊ะนะ",
    "age": "30",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "วีลแชร์ฟันดาบ",
    "sportSub": "เอเป้หญิง",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 60,
    "year": 2004,
    "city": "Athens",
    "name": "ประวัติ วะโฮรัมย์",
    "age": "23",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 10,000 เมตร",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 61,
    "year": 2004,
    "city": "Athens",
    "name": "สาคร ขันทะสิทธิ์, รัตนา เตชะมณีวัฒน์",
    "age": "25",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "วีลแชร์เทนนิส",
    "sportSub": "หญิงคู่",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 62,
    "year": 2004,
    "city": "Athens",
    "name": "วาสนา กาบไม้จันทร์",
    "age": "25",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ยิงธนู",
    "sportSub": "หญิงเดี่ยวท่ายืน",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 63,
    "year": 2004,
    "city": "Athens",
    "name": "ทองสา มารศรี",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "รุ่นไม่เกิน 48 กก.",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 64,
    "year": 2004,
    "city": "Athens",
    "name": "สมชาย ดวงแก้ว",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "กบ 50 เมตร",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 65,
    "year": 2004,
    "city": "Athens",
    "name": "สมชาย ดวงแก้ว",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ผีเสื้อ 50 เมตร",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 66,
    "year": 2004,
    "city": "Athens",
    "name": "พิเชษฐ์ กรุงเกตุ",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 200 เมตร",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 67,
    "year": 2004,
    "city": "Athens",
    "name": "เรวัติ ต๋านะ",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 10,000 เมตร",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 68,
    "year": 2004,
    "city": "Athens",
    "name": "สายสุนีย์ จ๊ะนะ",
    "age": "30",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "วีลแชร์ฟันดาบ",
    "sportSub": "ฟอยล์หญิง",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 69,
    "year": 2004,
    "city": "Athens",
    "name": "พัทยา เทศทอง",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "ชายเดี่ยว BC1",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 70,
    "year": 2004,
    "city": "Athens",
    "name": "พนม ลักษณะพริ้ม",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "กบ 100 เมตร",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 71,
    "year": 2004,
    "city": "Athens",
    "name": "สานิตย์ สงค์นอก",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "เดี่ยวผสม 150 เมตร",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 72,
    "year": 2008,
    "city": "Beijing",
    "name": "ประวัติ วะโฮรัมย์",
    "age": "27",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 5,000 เมตร",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 73,
    "year": 2008,
    "city": "Beijing",
    "name": "ประวัติ วะโฮรัมย์",
    "age": "27",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 1,500 เมตร",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 74,
    "year": 2008,
    "city": "Beijing",
    "name": "สายชล คนเจน",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 100 เมตร",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 75,
    "year": 2008,
    "city": "Beijing",
    "name": "สายชล คนเจน",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 200 เมตร",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 76,
    "year": 2008,
    "city": "Beijing",
    "name": "ประวัติ วะโฮรัมย์, สายชล คนเจน, พิเชษฐ์ กรุงเกตุ, ศุภชัย โกยทรัพย์",
    "age": "252",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 4×100 เมตร",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 77,
    "year": 2008,
    "city": "Beijing",
    "name": "ประวัติ วะโฮรัมย์, สายชล คนเจน, พิเชษฐ์ กรุงเกตุ, ศุภชัย โกยทรัพย์",
    "age": "252",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 4×400 เมตร",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 78,
    "year": 2008,
    "city": "Beijing",
    "name": "ประวัติ วะโฮรัมย์",
    "age": "27",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 800 เมตร",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 79,
    "year": 2008,
    "city": "Beijing",
    "name": "สายชล คนเจน",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 400 เมตร",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 80,
    "year": 2008,
    "city": "Beijing",
    "name": "ศุภชัย โกยทรัพย์",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 100 เมตร",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 81,
    "year": 2008,
    "city": "Beijing",
    "name": "เพ็ชร รุ่งศรี",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 200 เมตร",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 82,
    "year": 2008,
    "city": "Beijing",
    "name": "สมควร อนนท์",
    "age": "25",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "รุ่นไม่เกิน 52 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 83,
    "year": 2008,
    "city": "Beijing",
    "name": "ณรงค์ แคสนั่น",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "รุ่นไม่เกิน 52 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 84,
    "year": 2008,
    "city": "Beijing",
    "name": "สายสุนีย์ จ๊ะนะ",
    "age": "34",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "วีลแชร์ฟันดาบ",
    "sportSub": "เอเป้หญิง",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 85,
    "year": 2012,
    "city": "London",
    "name": "พัทธยา เทศทอง",
    "age": "33",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "เดี่ยว",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 86,
    "year": 2012,
    "city": "London",
    "name": "พัทธยา เทศทอง, วิษณุ ฮวดประดิษฐ์, มงคล จิตรเสงี่ยม, วัชรพล วงษา",
    "age": "252",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "ประเภททีม",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 87,
    "year": 2012,
    "city": "London",
    "name": "รุ่งโรจน์ ไทยนิยม",
    "age": "25",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ประเภทบุคคล",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 88,
    "year": 2012,
    "city": "London",
    "name": "สายสุนีย์ จ๊ะนะ",
    "age": "38",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "วีลแชร์ฟันดาบ",
    "sportSub": "ประเภทบุคคล",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 89,
    "year": 2012,
    "city": "London",
    "name": "ประวัติ วะโฮรัมย์",
    "age": "31",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 1,500 เมตร",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 90,
    "year": 2012,
    "city": "London",
    "name": "ประวัติ วะโฮรัมย์, สายชล คนเจน, โสภา อินทเสน, ศุภชัย โกยทรัพย์",
    "age": "252",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่งผลัด 4×400 เมตร",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 91,
    "year": 2012,
    "city": "London",
    "name": "สายชล คนเจน",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 100 เมตร",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 92,
    "year": 2012,
    "city": "London",
    "name": "สายชล คนเจน",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 800 เมตร",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 93,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "ประวัติ วะโฮรัมย์",
    "age": "35",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 5,000 เมตร T53/T54",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 94,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "ประวัติ วะโฮรัมย์",
    "age": "35",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 1,500 เมตร T54",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 95,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "พงศกร แปยอ",
    "age": "19",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 800 เมตร T52/T53",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 96,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "พงศกร แปยอ",
    "age": "19",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 400 เมตร T53",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 97,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "วัชรพล วงษา",
    "age": "25",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "บุคคล BC2",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 98,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "พัทธยา เทศทอง, วัชรพล วงษา, วรวุฒิ แสงอำภา, สุบิน ทิพย์มะณี",
    "age": "129",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "ทีม BC1/BC2",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 99,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "พงศกร แปยอ",
    "age": "19",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 100 เมตร T53",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 100,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "สายชล คนเจน",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 800 เมตร T54",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 101,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "พงศกร แปยอ, ประวัติ วะโฮรัมย์, เรวัตร์ ต๋านะ, สายชล คนเจน",
    "age": "252",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่งผลัด 4×400 เมตร T53/T54",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 102,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "วรวุฒิ แสงอำภา",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "บุคคล BC2",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 103,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "สายสุนีย์ จ๊ะนะ",
    "age": "42",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "วีลแชร์ฟันดาบ",
    "sportSub": "เอเป้บุคคลหญิง B",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 104,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "หาญฤชัย เนตศิริ",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ยิงธนู",
    "sportSub": "คันธนูโค้งกลับชาย",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 105,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "สายชล คนเจน",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 1,500 เมตร T53/T54",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 106,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "พิชญา คูรัตนศิริ",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 1,500 เมตร T51/T52",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 107,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "พรโชค ลาภเย็น",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "บุคคล BC4",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 108,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "เฉลิมพล ตันบุตร, พรโชค ลาภเย็น, นวลจันทร์ พลศิลา",
    "age": "129",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "คู่ BC4",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 109,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "รุ่งโรจน์ ไทยนิยม",
    "age": "29",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "บุคคล TT6",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 110,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "อนุรักษ์ ลาววงษ์, ยุทธจักร กลิ่นบานชื่น",
    "age": "252",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ทีมชาย TT3",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 111,
    "year": 2020,
    "city": "Tokyo",
    "name": "พงศกร แปยอ",
    "age": "24",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 400 เมตร T53",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 112,
    "year": 2020,
    "city": "Tokyo",
    "name": "พงศกร แปยอ",
    "age": "24",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 100 เมตร T53",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 113,
    "year": 2020,
    "city": "Tokyo",
    "name": "พงศกร แปยอ",
    "age": "24",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 800 เมตร T53",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 114,
    "year": 2020,
    "city": "Tokyo",
    "name": "อธิวัฒน์ แพงเหนือ",
    "age": "18",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 100 เมตร T54",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 115,
    "year": 2020,
    "city": "Tokyo",
    "name": "วิษณุ ฮวดประดิษฐ์, สุบิน ทิพย์มะนี, วัชรพล วงษา, วรวุฒิ แสงอำภา",
    "age": "129",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "ทีม BC1–2",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 116,
    "year": 2020,
    "city": "Tokyo",
    "name": "อธิวัฒน์ แพงเหนือ",
    "age": "18",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 400 เมตร T54",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 117,
    "year": 2020,
    "city": "Tokyo",
    "name": "ประวัติ วะโฮรัมย์",
    "age": "40",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 1,500 เมตร T54",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 118,
    "year": 2020,
    "city": "Tokyo",
    "name": "พรโชค ลาภเย็น",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "บุคคล BC4",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 119,
    "year": 2020,
    "city": "Tokyo",
    "name": "วัชรพล วงษา",
    "age": "30",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "บุคคล BC2",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 120,
    "year": 2020,
    "city": "Tokyo",
    "name": "สุจิรัตน์ ปุกคำ",
    "age": "35",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "หญิงเดี่ยว WH1",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 121,
    "year": 2020,
    "city": "Tokyo",
    "name": "สายสุนีย์ จ๊ะนะ",
    "age": "47",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "วีลแชร์ฟันดาบ",
    "sportSub": "เอเป้ คลาส B",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 122,
    "year": 2020,
    "city": "Tokyo",
    "name": "รุ่งโรจน์ ไทยนิยม",
    "age": "34",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "บุคคลชาย คลาส 6",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 123,
    "year": 2020,
    "city": "Tokyo",
    "name": "ภูธเรศ คงรักษ์",
    "age": "26",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 5,000 เมตร T54",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 124,
    "year": 2020,
    "city": "Tokyo",
    "name": "ภูธเรศ คงรักษ์",
    "age": "26",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 1,500 เมตร T54",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 125,
    "year": 2020,
    "city": "Tokyo",
    "name": "สายชล คนเจน",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 800 เมตร T54",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 126,
    "year": 2020,
    "city": "Tokyo",
    "name": "อนุรักษ์ ลาววงษ์, ยุทธจักร กลิ่นบานชื่น, ถิรายุ เชื้อวงษ์",
    "age": "252",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ทีมชาย คลาส 3",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 127,
    "year": 2020,
    "city": "Tokyo",
    "name": "ขวัญสุดา พวงกิจจา",
    "age": "20",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "เทควันโด",
    "sportSub": "รุ่นไม่เกิน 49 กก. K44",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 128,
    "year": 2020,
    "city": "Tokyo",
    "name": "สุจิรัตน์ ปุกคำ, อำนวย เวชวิฐาน",
    "age": "25",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "หญิงคู่ WH1–WH2",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 129,
    "year": 2024,
    "city": "Paris",
    "name": "สายสุนีย์ จ๊ะนะ",
    "age": "50",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "วีลแชร์ฟันดาบ",
    "sportSub": "เซเบอร์บุคคลหญิง คลาส B",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 130,
    "year": 2024,
    "city": "Paris",
    "name": "สายสุนีย์ จ๊ะนะ",
    "age": "50",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "วีลแชร์ฟันดาบ",
    "sportSub": "ฟอยล์บุคคลหญิง คลาส B",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 131,
    "year": 2024,
    "city": "Paris",
    "name": "สายสุนีย์ จ๊ะนะ",
    "age": "50",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "วีลแชร์ฟันดาบ",
    "sportSub": "เอเป้บุคคลหญิง คลาส B",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 132,
    "year": 2024,
    "city": "Paris",
    "name": "พงศกร แปยอ",
    "age": "27",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 400 เมตร T53",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 133,
    "year": 2024,
    "city": "Paris",
    "name": "ชัยวัฒน์ รัตนะ",
    "age": "28",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 100 เมตร T34",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 134,
    "year": 2024,
    "city": "Paris",
    "name": "วรวุฒิ แสงอำภา",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "บุคคล BC2",
    "medal": "ทอง",
    "countryBorn": "Thailand"
  },
  {
    "number": 135,
    "year": 2024,
    "city": "Paris",
    "name": "รุ่งโรจน์ ไทยนิยม, พิสิษฐ์ หวังผลพัฒนศิริ",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ชายคู่ MD14",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 136,
    "year": 2024,
    "city": "Paris",
    "name": "ยุทธจักร กลิ่นบานชื่น, วิจิตรา ใจอ่อน",
    "age": "129",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "คู่ผสม MD7",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 137,
    "year": 2024,
    "city": "Paris",
    "name": "รุ่งโรจน์ ไทยนิยม",
    "age": "37",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ชายเดี่ยว MS6",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 138,
    "year": 2024,
    "city": "Paris",
    "name": "วันชัย ชัยวุฒิ",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ชายเดี่ยว MS4",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 139,
    "year": 2024,
    "city": "Paris",
    "name": "พงศกร แปยอ",
    "age": "27",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 100 เมตร T53",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 140,
    "year": 2024,
    "city": "Paris",
    "name": "พงศกร แปยอ",
    "age": "27",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 800 เมตร T53",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 141,
    "year": 2024,
    "city": "Paris",
    "name": "อธิวัฒน์ แพงเหนือ",
    "age": "21",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 100 เมตร T54",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 142,
    "year": 2024,
    "city": "Paris",
    "name": "อธิวัฒน์ แพงเหนือ",
    "age": "21",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 400 เมตร T54",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 143,
    "year": 2024,
    "city": "Paris",
    "name": "ชัยวัฒน์ รัตนะ",
    "age": "28",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วีลแชร์เรซซิ่ง 800 เมตร T34",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 144,
    "year": 2024,
    "city": "Paris",
    "name": "สุจิรัตน์ ปุกคำ",
    "age": "38",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "หญิงเดี่ยว WH",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 145,
    "year": 2024,
    "city": "Paris",
    "name": "วิสิทธิ์ กิ่งมะนาว",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "วีลแชร์ฟันดาบ",
    "sportSub": "เอเป้บุคคลชาย คลาส B",
    "medal": "เงิน",
    "countryBorn": "Thailand"
  },
  {
    "number": 146,
    "year": 2024,
    "city": "Paris",
    "name": "ขวัญสุดา พวงกิจจา",
    "age": "23",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "เทควันโด",
    "sportSub": "รุ่น 47 กก. K44",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 147,
    "year": 2024,
    "city": "Paris",
    "name": "วันชัย ชัยวุฒิ, ยุทธจักร กลิ่นบานชื่น",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ชายคู่ MD8",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 148,
    "year": 2024,
    "city": "Paris",
    "name": "ดารารัตน์ อาสายุทธ์, ชิลชิตพยัค บุตรวรรณสิริณา",
    "age": "25",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "หญิงคู่ WD5",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 149,
    "year": 2024,
    "city": "Paris",
    "name": "ยุทธจักร กลิ่นบานชื่น",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ชายเดี่ยว MS3",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 150,
    "year": 2024,
    "city": "Paris",
    "name": "เฉลิมพงษ์ พันภู่",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ชายเดี่ยว MS7",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 151,
    "year": 2024,
    "city": "Paris",
    "name": "พิสิษฐ์ หวังผลพัฒนศิริ",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ชายเดี่ยว MS8",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 152,
    "year": 2024,
    "city": "Paris",
    "name": "สุจิรัตน์ ปุกคำ, อำนวย เวชวิฐาน",
    "age": "25",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "หญิงคู่ WH1–WH2",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 153,
    "year": 2024,
    "city": "Paris",
    "name": "มงคล บุญสุน",
    "age": "29",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "ชายเดี่ยว SL3",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 154,
    "year": 2024,
    "city": "Paris",
    "name": "วัชรพล วงษา",
    "age": "33",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "บุคคล BC2",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 155,
    "year": 2024,
    "city": "Paris",
    "name": "พรโชค ลาภเย็น, นวลจันทร์ พลศิลา",
    "age": "129",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "คู่ผสม BC4",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 156,
    "year": 2024,
    "city": "Paris",
    "name": "กมลพรรณ กระราชเพชร",
    "age": "25",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "รุ่นไม่เกิน 55 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 157,
    "year": 2024,
    "city": "Paris",
    "name": "ศศิราวรรณ อินทโชติ",
    "age": "25",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วิ่ง 200 เมตร T47",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 158,
    "year": 2024,
    "city": "Paris",
    "name": "สายสุนีย์ จ๊ะนะ, เดือน นาคประสิทธิ์, อภิญญา ทองแดง",
    "age": "25",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "วีลแชร์ฟันดาบ",
    "sportSub": "เอเป้ทีมหญิง",
    "medal": "ทองแดง",
    "countryBorn": "Thailand"
  },
  {
    "number": 159,
    "year": 1988,
    "city": "Seoul",
    "name": "Lilies Handayani, Nurfitriyana Saiman, Kusuma Wardhani",
    "age": "23, 26, 24",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยิงธนู",
    "sportSub": "ทีมหญิง",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 160,
    "year": 1992,
    "city": "Barcelona",
    "name": "Alan Budikusuma",
    "age": "24",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 161,
    "year": 1992,
    "city": "Barcelona",
    "name": "Susy Susanti",
    "age": "21",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทหญิงเดี่ยว",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 162,
    "year": 1992,
    "city": "Barcelona",
    "name": "Ardy Wiranata",
    "age": "22",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 163,
    "year": 1992,
    "city": "Barcelona",
    "name": "Eddy Hartono, Rudy Gunawan",
    "age": "28, 26",
    "gender": "ทีมชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายคู่",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 164,
    "year": 1992,
    "city": "Barcelona",
    "name": "Hermawan Susanto",
    "age": "25",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 165,
    "year": 1996,
    "city": "Atlanta",
    "name": "Rexy Mainaky, Ricky Subagja",
    "age": "28, 25",
    "gender": "ทีมชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายคู่",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 166,
    "year": 1996,
    "city": "Atlanta",
    "name": "Mia Audina",
    "age": "17",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทหญิงเดี่ยว",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 167,
    "year": 1996,
    "city": "Atlanta",
    "name": "Antonius Ariantho, Denny Kantono",
    "age": "23, 26",
    "gender": "ทีมชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายคู่",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 168,
    "year": 1996,
    "city": "Atlanta",
    "name": "Susy Susanti",
    "age": "25",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทหญิงเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 169,
    "year": 2000,
    "city": "Sydney",
    "name": "Candra Wijaya, Tony Gunawan",
    "age": "25, 25",
    "gender": "ทีมชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายคู่",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 170,
    "year": 2000,
    "city": "Sydney",
    "name": "Hendrawan",
    "age": "28",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 171,
    "year": 2000,
    "city": "Sydney",
    "name": "Minarti Timur, Tri Kusharyanto",
    "age": "32, 26",
    "gender": "ทีมผสม",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "คู่ผสม",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 172,
    "year": 2000,
    "city": "Sydney",
    "name": "Raema Lisa Rumbewas",
    "age": "20",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง น้ำหนัก 48 กก.",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 173,
    "year": 2000,
    "city": "Sydney",
    "name": "Sri Indriyani",
    "age": "22",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง น้ำหนัก 48 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 174,
    "year": 2000,
    "city": "Sydney",
    "name": "Winarni Binti Slamet",
    "age": "25",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง น้ำหนัก 53 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 175,
    "year": 2004,
    "city": "Athens",
    "name": "Taufik Hidayat",
    "age": "23",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 176,
    "year": 2004,
    "city": "Athens",
    "name": "Sony Dwi Kuncoro",
    "age": "20",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 177,
    "year": 2004,
    "city": "Athens",
    "name": "Eng Hian, Flandy Limpele",
    "age": "27, 30",
    "gender": "ทีมชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายคู่",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 178,
    "year": 2004,
    "city": "Athens",
    "name": "Raema Lisa Rumbewas",
    "age": "24",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง น้ำหนัก 53 กก.",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 179,
    "year": 2008,
    "city": "Beijing",
    "name": "Markis Kido,Hendra Setiawan",
    "age": "24",
    "gender": "ทีมชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายคู่",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 180,
    "year": 2008,
    "city": "Beijing",
    "name": "Maria Yulianti",
    "age": "23",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทหญิงเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 181,
    "year": 2008,
    "city": "Beijing",
    "name": "Liliyana Natsir, Nova Widianto",
    "age": "23, 31",
    "gender": "ทีมผสม",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "คู่ผสม",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 182,
    "year": 2008,
    "city": "Beijing",
    "name": "Eko Yuli Irawan",
    "age": "19",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 56 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 183,
    "year": 2008,
    "city": "Beijing",
    "name": "Triyatno",
    "age": "21",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 62 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 184,
    "year": 2008,
    "city": "Beijing",
    "name": "Raema Lisa Rumbewas",
    "age": "28",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง น้ำหนัก 53 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 185,
    "year": 2012,
    "city": "London",
    "name": "Eko Yuli Irawan",
    "age": "23",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 62 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 186,
    "year": 2012,
    "city": "London",
    "name": "Triyatno",
    "age": "25",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 69 กก.",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 187,
    "year": 2012,
    "city": "London",
    "name": "Citra Febrianti",
    "age": "24",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง น้ำหนัก 53 กก.",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 188,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Tontowi Ahmad, Liliyana Natsir",
    "age": "29, 31",
    "gender": "ทีมผสม",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "คู่ผสม",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 189,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Eko Yuli Irawan",
    "age": "27",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 62 กก.",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 190,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Sri Wahyuni Agustiani",
    "age": "22",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง น้ำหนัก 48 กก.",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 191,
    "year": 2020,
    "city": "Tokyo",
    "name": "Anthony Sinisuka Ginting",
    "age": "24",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 192,
    "year": 2020,
    "city": "Tokyo",
    "name": "Greysia Polii, Apriyani Rahayu",
    "age": "24, 22",
    "gender": "ทีมหญิง",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทหญิงคู่",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 193,
    "year": 2020,
    "city": "Tokyo",
    "name": "Eko Yuli Irawan",
    "age": "31",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 61 กก.",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 194,
    "year": 2020,
    "city": "Tokyo",
    "name": "Rahmat Abdullah",
    "age": "20",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 73 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 195,
    "year": 2020,
    "city": "Tokyo",
    "name": "Windy Cantika Aisah",
    "age": "18",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง น้ำหนัก 49 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 196,
    "year": 2024,
    "city": "Paris",
    "name": "Rizki Juniansyah",
    "age": "21",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 73 กก.",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 197,
    "year": 2024,
    "city": "Paris",
    "name": "Veddriq Leonardo",
    "age": "27",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ปีนผา",
    "sportSub": "ความเร็ว ชาย",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 198,
    "year": 2024,
    "city": "Paris",
    "name": "Gregoria Mariska Tunjung",
    "age": "25",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทหญิงเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 199,
    "year": 1976,
    "city": "montreal",
    "name": "Itria Dini",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "พุ่งแหลนแม่นยำชาย",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 200,
    "year": 1976,
    "city": "montreal",
    "name": "Syarifuddin",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ลอนโบว์ลิ่ง",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 201,
    "year": 1976,
    "city": "montreal",
    "name": "Ashari",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วิ่ง 100 เมตรชาย",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 202,
    "year": 1976,
    "city": "montreal",
    "name": "Itria Dini",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "ทุ่มน้ำหนักชาย",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 203,
    "year": 1976,
    "city": "montreal",
    "name": "Saneng Hanafi",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "ขว้างจักรชาย",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 204,
    "year": 1976,
    "city": "montreal",
    "name": "Saneng Hanafi",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "พุ่งแหลนชาย",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 205,
    "year": 1980,
    "city": "Arnhem",
    "name": "Yan Soebiyanto",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ลอนโบว์ลิ่ง",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 206,
    "year": 1980,
    "city": "Arnhem",
    "name": "R.S. Arlen",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชายรุ่นไม่เกิน 57 กก.",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 207,
    "year": 1980,
    "city": "Arnhem",
    "name": "Sigit Soepadi",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ลอนโบว์ลิ่ง",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 208,
    "year": 1980,
    "city": "Arnhem",
    "name": "Soekarsan",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ลอนโบว์ลิ่ง",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 209,
    "year": 1980,
    "city": "Arnhem",
    "name": "Moenali, Yamin Ismail",
    "age": "252",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ลอนโบว์ลิ่ง",
    "sportSub": "ประเภทคู่ชาย",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 210,
    "year": 1980,
    "city": "Arnhem",
    "name": "R.S. Arlen, Safri Tanjung",
    "age": "252",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ลอนโบว์ลิ่ง",
    "sportSub": "ประเภทคู่ชาย",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 211,
    "year": 1984,
    "city": "New York/Stoke Mandeville",
    "name": "Ninik Umardiyani",
    "age": "25",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ลอนโบว์ลิ่ง",
    "sportSub": "ประเภทหญิงเดี่ยว",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 212,
    "year": 1984,
    "city": "New York/Stoke Mandeville",
    "name": "Kurnianto, Memed Lesmana",
    "age": "252",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "ลอนโบว์ลิ่ง",
    "sportSub": "ประเภทคู่ชาย",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 213,
    "year": 1988,
    "city": "Seoul",
    "name": "Hadi Abdulaziz",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "กระโดดสูงชาย",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 214,
    "year": 1988,
    "city": "Seoul",
    "name": "Soeparni",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "ทุ่มน้ำหนักชาย",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 215,
    "year": 2012,
    "city": "London",
    "name": "David Jacobs",
    "age": "35",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 216,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Ni Nengah Widiasih",
    "age": "23",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง น้ำหนัก 41 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 217,
    "year": 2020,
    "city": "Tokyo",
    "name": "Leani Ratri Oktila, Khalimatus Sadiyah",
    "age": "30,21",
    "gender": "ทีมหญิง",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทหญิงคู่",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 218,
    "year": 2020,
    "city": "Tokyo",
    "name": "Hary Susanto, Leani Ratri Oktila",
    "age": "46,30",
    "gender": "ทีมผสม",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "คู่ผสม",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 219,
    "year": 2020,
    "city": "Tokyo",
    "name": "Ni Nengah Widiasih",
    "age": "28",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง น้ำหนัก 41 กก.",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 220,
    "year": 2020,
    "city": "Tokyo",
    "name": "Dheva Anrimusthi",
    "age": "22",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 221,
    "year": 2020,
    "city": "Tokyo",
    "name": "Leani Ratri Oktila",
    "age": "30",
    "gender": "ทีมหญิง",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทหญิงคู่",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 222,
    "year": 2020,
    "city": "Tokyo",
    "name": "Saptoyogo Purnomo",
    "age": "22",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วิ่ง 100 เมตรชาย",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 223,
    "year": 2020,
    "city": "Tokyo",
    "name": "David Jacobs",
    "age": "44",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 224,
    "year": 2020,
    "city": "Tokyo",
    "name": "Suryo Nugroho",
    "age": "24",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 225,
    "year": 2020,
    "city": "Tokyo",
    "name": "Fredy Setiawan",
    "age": "23",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 226,
    "year": 2024,
    "city": "Paris",
    "name": "Hikmat Ramdani, Leani Ratri Oktila",
    "age": "20,33",
    "gender": "ทีมผสม",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "คู่ผสม",
    "medal": "ทอง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 227,
    "year": 2024,
    "city": "Paris",
    "name": "Saptoyogo Purnomo",
    "age": "25",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "100 เมตรชาย",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 228,
    "year": 2024,
    "city": "Paris",
    "name": "Muhammad Bintang, Herlangga",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 229,
    "year": 2024,
    "city": "Paris",
    "name": "Qonitah Ikhtiar, Syakuroh",
    "age": "25",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทหญิงเดี่ยว",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 230,
    "year": 2024,
    "city": "Paris",
    "name": "Fredy Setiawan, Khalimatus Sadiyah",
    "age": "26,24",
    "gender": "ทีมผสม",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "คู่ผสม",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 231,
    "year": 2024,
    "city": "Paris",
    "name": "Leani Ratri Oktila",
    "age": "33",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทหญิงเดี่ยว",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 232,
    "year": 2024,
    "city": "Paris",
    "name": "Suryo Nugroho",
    "age": "31",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 233,
    "year": 2024,
    "city": "Paris",
    "name": "Muhamad Afrizal Syafa, Felix Ardi Yudha, Gischa Zayana",
    "age": "129",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "คู่ผสม BC1–2",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 234,
    "year": 2024,
    "city": "Paris",
    "name": "Karisma Evi Tiarani",
    "age": "23",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วิ่ง 100 เมตรหญิง",
    "medal": "เงิน",
    "countryBorn": "Indonesia"
  },
  {
    "number": 235,
    "year": 2024,
    "city": "Paris",
    "name": "Gischa Zayana",
    "age": "25",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "ว่ายน้ำหญิงเดี่ยว BC2",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 236,
    "year": 2024,
    "city": "Paris",
    "name": "Muhamad Afrizal Syafa",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "ประเภทชายเดี่ยว BC1",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 237,
    "year": 2024,
    "city": "Paris",
    "name": "Subhan, Rina Marlina",
    "age": "32,33",
    "gender": "ทีมผสม",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "คู่ผสม",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 238,
    "year": 2024,
    "city": "Paris",
    "name": "Dheva Anrimusthi",
    "age": "25",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 239,
    "year": 2024,
    "city": "Paris",
    "name": "Fredy Setiawan",
    "age": "26",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Indonesia"
  },
  {
    "number": 240,
    "year": 1928,
    "city": "Amsterdam",
    "name": "Teófilo Yldefonso",
    "age": "25",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ท่ากบ 200 เมตรชาย",
    "medal": "ทองแดง",
    "countryBorn": "Philippines"
  },
  {
    "number": 241,
    "year": 1932,
    "city": "Los Angeles",
    "name": "Simeon Toribio",
    "age": "26",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "กรีฑา",
    "sportSub": "กระโดดสูงชาย",
    "medal": "ทองแดง",
    "countryBorn": "Philippines"
  },
  {
    "number": 242,
    "year": 1932,
    "city": "Los Angeles",
    "name": "Teófilo Yldefonso",
    "age": "29",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ท่ากบ 200 เมตรชาย",
    "medal": "ทองแดง",
    "countryBorn": "Philippines"
  },
  {
    "number": 243,
    "year": 1932,
    "city": "Los Angeles",
    "name": "José Villanueva",
    "age": "19",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นแบนตัมเวต",
    "medal": "ทองแดง",
    "countryBorn": "Philippines"
  },
  {
    "number": 244,
    "year": 1936,
    "city": "Berlin",
    "name": "Miguel White",
    "age": "26",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "กรีฑา",
    "sportSub": "วิ่งข้ามรั้ว 400 เมตรชาย",
    "medal": "ทองแดง",
    "countryBorn": "Philippines"
  },
  {
    "number": 245,
    "year": 1964,
    "city": "Tokyo",
    "name": "Anthony Villanueva",
    "age": "19",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นเฟเธอร์เวต",
    "medal": "เงิน",
    "countryBorn": "Philippines"
  },
  {
    "number": 246,
    "year": 1988,
    "city": "Seoul",
    "name": "Leopoldo Serantes",
    "age": "26",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นไลต์ฟลายเวต",
    "medal": "ทองแดง",
    "countryBorn": "Philippines"
  },
  {
    "number": 247,
    "year": 1992,
    "city": "Barcelona",
    "name": "Roel Velasco",
    "age": "20",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นไลต์ฟลายเวต",
    "medal": "ทองแดง",
    "countryBorn": "Philippines"
  },
  {
    "number": 248,
    "year": 1996,
    "city": "Atlanta",
    "name": "Onyok Velasco",
    "age": "22",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นไลต์ฟลายเวต",
    "medal": "เงิน",
    "countryBorn": "Philippines"
  },
  {
    "number": 249,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Hidilyn Diaz",
    "age": "25",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง น้ำหนัก 53 กก.",
    "medal": "เงิน",
    "countryBorn": "Philippines"
  },
  {
    "number": 250,
    "year": 2020,
    "city": "Tokyo",
    "name": "Hidilyn Diaz",
    "age": "30",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง น้ำหนัก 55 กก.",
    "medal": "ทอง",
    "countryBorn": "Philippines"
  },
  {
    "number": 251,
    "year": 2020,
    "city": "Tokyo",
    "name": "Nesthy Petecio",
    "age": "29",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นเฟเธอร์เวต",
    "medal": "เงิน",
    "countryBorn": "Philippines"
  },
  {
    "number": 252,
    "year": 2020,
    "city": "Tokyo",
    "name": "Carlo Paalam",
    "age": "23",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นไลต์ฟลายเวต",
    "medal": "เงิน",
    "countryBorn": "Philippines"
  },
  {
    "number": 253,
    "year": 2020,
    "city": "Tokyo",
    "name": "Eumir Marcial",
    "age": "25",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นมิดเดิลเวต",
    "medal": "ทองแดง",
    "countryBorn": "Philippines"
  },
  {
    "number": 254,
    "year": 2024,
    "city": "Paris",
    "name": "Carlos Yulo",
    "age": "24",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยิมนาสติก",
    "sportSub": "กายกรรมพื้นชาย",
    "medal": "ทอง",
    "countryBorn": "Philippines"
  },
  {
    "number": 255,
    "year": 2024,
    "city": "Paris",
    "name": "Carlos Yulo",
    "age": "24",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยิมนาสติก",
    "sportSub": "กระโดดข้ามม้าชาย",
    "medal": "ทอง",
    "countryBorn": "Philippines"
  },
  {
    "number": 256,
    "year": 2024,
    "city": "Paris",
    "name": "Aira Villegas",
    "age": "29",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นไลต์ฟลายเวต",
    "medal": "ทองแดง",
    "countryBorn": "Philippines"
  },
  {
    "number": 257,
    "year": 2024,
    "city": "Paris",
    "name": "Nesthy Petecio",
    "age": "32",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "มวยสากลสมัครเล่น",
    "sportSub": "รุ่นเฟเธอร์เวต",
    "medal": "ทองแดง",
    "countryBorn": "Philippines"
  },
  {
    "number": 258,
    "year": 2000,
    "city": "Sydney",
    "name": "Adeline Dumapong",
    "age": "35",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง น้ำหนัก 82.5 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Philippines"
  },
  {
    "number": 259,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Josephine Medina",
    "age": "46",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ประเภทหญิงเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Philippines"
  },
  {
    "number": 260,
    "year": 1992,
    "city": "Barcelona",
    "name": "Razif Sidek, Jalani Sidek",
    "age": "30,28",
    "gender": "ทีมชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทคู่ชาย",
    "medal": "ทองแดง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 261,
    "year": 1996,
    "city": "Atlanta",
    "name": "Cheah Soon Kit, Yap Kim Hock",
    "age": "26,20",
    "gender": "ทีมชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ชายคู่",
    "medal": "เงิน",
    "countryBorn": "Malaysia"
  },
  {
    "number": 262,
    "year": 1996,
    "city": "Atlanta",
    "name": "Rashid Sidek",
    "age": "28",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ชายเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 263,
    "year": 2008,
    "city": "Beijing",
    "name": "Lee Chong Wei",
    "age": "26",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ชายเดี่ยว",
    "medal": "เงิน",
    "countryBorn": "Malaysia"
  },
  {
    "number": 264,
    "year": 2012,
    "city": "London",
    "name": "Lee Chong Wei",
    "age": "30",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ชายเดี่ยว",
    "medal": "เงิน",
    "countryBorn": "Malaysia"
  },
  {
    "number": 265,
    "year": 2012,
    "city": "London",
    "name": "Pandelela Rinong",
    "age": "19",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "กระโดดน้ำ",
    "sportSub": "แพลตฟอร์ม 10 เมตร หญิง",
    "medal": "ทองแดง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 266,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Lee Chong Wei",
    "age": "34",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ชายเดี่ยว",
    "medal": "เงิน",
    "countryBorn": "Malaysia"
  },
  {
    "number": 267,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Goh V Shem, Tan Wee Kiong",
    "age": "27,22",
    "gender": "ทีมชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ชายคู่",
    "medal": "เงิน",
    "countryBorn": "Malaysia"
  },
  {
    "number": 268,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Chan Peng Soon, Goh Liu Ying",
    "age": "28,27",
    "gender": "ทีมผสม",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "คู่ผสม",
    "medal": "เงิน",
    "countryBorn": "Malaysia"
  },
  {
    "number": 269,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Cheong Jun Hoong, Pandelela Rinong",
    "age": "27,19",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "กระโดดน้ำ",
    "sportSub": "แพลตฟอร์มซิงโครไนซ์ 10 เมตร หญิง",
    "medal": "เงิน",
    "countryBorn": "Malaysia"
  },
  {
    "number": 270,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Azizulhasni Awang",
    "age": "28",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "จักรยาน",
    "sportSub": "คีรินชาย",
    "medal": "ทองแดง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 271,
    "year": 2020,
    "city": "Tokyo",
    "name": "Azizulhasni Awang",
    "age": "33",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "จักรยาน",
    "sportSub": "คีรินชาย",
    "medal": "เงิน",
    "countryBorn": "Malaysia"
  },
  {
    "number": 272,
    "year": 2020,
    "city": "Tokyo",
    "name": "Aaron Chia, Soh Wooi Yik",
    "age": "24,27",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ชายคู่",
    "medal": "ทองแดง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 273,
    "year": 2024,
    "city": "Paris",
    "name": "Aaron Chia, Soh Wooi Yik",
    "age": "26, 23",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ชายคู่",
    "medal": "ทองแดง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 274,
    "year": 2024,
    "city": "Paris",
    "name": "Lee Zii Jia",
    "age": "26",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แบดมินตัน",
    "sportSub": "ชายเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 275,
    "year": 1988,
    "city": "Seoul",
    "name": "Mariappan Perumal",
    "age": "20",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 57 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 276,
    "year": 1992,
    "city": "Barcelona",
    "name": "Cheok Kon Fatt",
    "age": "24",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 52 กก.",
    "medal": "เงิน",
    "countryBorn": "Malaysia"
  },
  {
    "number": 277,
    "year": 1992,
    "city": "Barcelona",
    "name": "Mariappan Perumal",
    "age": "24",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 60 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 278,
    "year": 1992,
    "city": "Barcelona",
    "name": "Mohamad Khasseri Othman",
    "age": "23",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "กระโดดสูงชาย",
    "medal": "ทองแดง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 279,
    "year": 2008,
    "city": "Beijing",
    "name": "Siow Lee Chan",
    "age": "35",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง น้ำหนัก 56 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 280,
    "year": 2012,
    "city": "London",
    "name": "Hasihin Sanawi",
    "age": "34",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "ยิงธนู",
    "sportSub": "ธนูโค้งเดี่ยวชาย",
    "medal": "เงิน",
    "countryBorn": "Malaysia"
  },
  {
    "number": 281,
    "year": 2012,
    "city": "London",
    "name": "Muhammad Ziyad Zolkefli",
    "age": "32",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "ทุ่มน้ำหนักชาย",
    "medal": "ทองแดง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 282,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Mohamad Ridzuan Mohamad Puzi",
    "age": "29",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วิ่ง 100 เมตรชาย",
    "medal": "ทอง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 283,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Muhammad Ziyad Zolkefli",
    "age": "32",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "ทุ่มน้ำหนักชาย",
    "medal": "ทอง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 284,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Abdul Latif Romly",
    "age": "29",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "กระโดดไกลชาย",
    "medal": "ทอง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 285,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Siti Noor Radiah Ismail",
    "age": "28",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "กระโดดไกลหญิง",
    "medal": "ทองแดง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 286,
    "year": 2020,
    "city": "Tokyo",
    "name": "Abdul Latif Romly",
    "age": "34",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "กระโดดไกลชาย",
    "medal": "ทอง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 287,
    "year": 2020,
    "city": "Tokyo",
    "name": "Cheah Liek Hou",
    "age": "33",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทอง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 288,
    "year": 2020,
    "city": "Tokyo",
    "name": "Bonnie Bunyau Gustin",
    "age": "35",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 72 กก.",
    "medal": "ทอง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 289,
    "year": 2020,
    "city": "Tokyo",
    "name": "Jong Yee Khie",
    "age": "32",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 107 กก.",
    "medal": "เงิน",
    "countryBorn": "Malaysia"
  },
  {
    "number": 290,
    "year": 2020,
    "city": "Tokyo",
    "name": "Chew Wei Lun",
    "age": "28",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "ประเภทผสม",
    "medal": "เงิน",
    "countryBorn": "Malaysia"
  },
  {
    "number": 291,
    "year": 2024,
    "city": "Paris",
    "name": "Cheah Liek Hou",
    "age": "37",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "แบดมินตัน",
    "sportSub": "ประเภทชายเดี่ยว",
    "medal": "ทอง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 292,
    "year": 2024,
    "city": "Paris",
    "name": "Bonnie Bunyau Gustin",
    "age": "39",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 72 กก.",
    "medal": "ทอง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 293,
    "year": 2024,
    "city": "Paris",
    "name": "Muhammad Ziyad Zolkefli",
    "age": "35",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "ทุ่มน้ำหนักชาย",
    "medal": "เงิน",
    "countryBorn": "Malaysia"
  },
  {
    "number": 294,
    "year": 2024,
    "city": "Paris",
    "name": "Abdul Latif Romly",
    "age": "27",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "กระโดดไกลชาย",
    "medal": "เงิน",
    "countryBorn": "Malaysia"
  },
  {
    "number": 295,
    "year": 2024,
    "city": "Paris",
    "name": "Eddy Bernard",
    "age": "27",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "100 เมตรชาย",
    "medal": "ทองแดง",
    "countryBorn": "Malaysia"
  },
  {
    "number": 296,
    "year": 1960,
    "city": "Rome",
    "name": "Tan Howe Liang",
    "age": "27",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "Lightweight",
    "medal": "เงิน",
    "countryBorn": "Singapore"
  },
  {
    "number": 297,
    "year": 2008,
    "city": "Beijing",
    "name": "Feng Tianwei, Li Jiawei, Wang Yuegu",
    "age": "21, 27, 27",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ทีมหญิง",
    "medal": "เงิน",
    "countryBorn": "Singapore"
  },
  {
    "number": 298,
    "year": 2012,
    "city": "London",
    "name": "Feng Tianwei",
    "age": "25",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ประเภทหญิงเดี่ยว",
    "medal": "ทองแดง",
    "countryBorn": "Singapore"
  },
  {
    "number": 299,
    "year": 2012,
    "city": "London",
    "name": "Feng Tianwei, Li Jiawei, Wang Yuegu",
    "age": "25, 30, 32",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "เทเบิลเทนนิส",
    "sportSub": "ทีมหญิง",
    "medal": "ทองแดง",
    "countryBorn": "Singapore"
  },
  {
    "number": 300,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Joseph Schooling",
    "age": "21",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ท่าผีเสื้อ 100 เมตรชาย",
    "medal": "ทอง",
    "countryBorn": "Singapore"
  },
  {
    "number": 301,
    "year": 2024,
    "city": "Paris",
    "name": "Max Maeder",
    "age": "17",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "แล่นเรือใบ",
    "sportSub": "ฟอร์มูล่าไคท์ ชาย",
    "medal": "ทองแดง",
    "countryBorn": "Singapore"
  },
  {
    "number": 302,
    "year": 2008,
    "city": "Beijing",
    "name": "Yip Pin Xiu",
    "age": "16",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ท่ากบ 50 เมตรหญิง",
    "medal": "ทอง",
    "countryBorn": "Singapore"
  },
  {
    "number": 303,
    "year": 2008,
    "city": "Beijing",
    "name": "Yip Pin Xiu",
    "age": "16",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ฟรีสไตล์หญิง 50 เมตร",
    "medal": "เงิน",
    "countryBorn": "Singapore"
  },
  {
    "number": 304,
    "year": 2008,
    "city": "Beijing",
    "name": "Laurentia Tan",
    "age": "29",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ขี่ม้า",
    "sportSub": "ทดสอบชิงแชมป์รายบุคคล ระดับ Ia",
    "medal": "ทองแดง",
    "countryBorn": "Singapore"
  },
  {
    "number": 305,
    "year": 2008,
    "city": "Beijing",
    "name": "Laurentia Tan",
    "age": "29",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ขี่ม้า",
    "sportSub": "ทดสอบฟรีสไตล์รายบุคคล ระดับ Ia",
    "medal": "ทองแดง",
    "countryBorn": "Singapore"
  },
  {
    "number": 306,
    "year": 2012,
    "city": "London",
    "name": "Laurentia Tan",
    "age": "33",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ขี่ม้า",
    "sportSub": "ทดสอบฟรีสไตล์รายบุคคล ระดับ Ia",
    "medal": "เงิน",
    "countryBorn": "Singapore"
  },
  {
    "number": 307,
    "year": 2012,
    "city": "London",
    "name": "Laurentia Tan",
    "age": "33",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ขี่ม้า",
    "sportSub": "ทดสอบชิงแชมป์รายบุคคล ระดับ Ia",
    "medal": "ทองแดง",
    "countryBorn": "Singapore"
  },
  {
    "number": 308,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Yip Pin Xiu",
    "age": "24",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ท่ากบ 100 เมตรหญิง",
    "medal": "ทอง",
    "countryBorn": "Singapore"
  },
  {
    "number": 309,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Yip Pin Xiu",
    "age": "24",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ท่ากบ 50 เมตรหญิง",
    "medal": "ทอง",
    "countryBorn": "Singapore"
  },
  {
    "number": 310,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Theresa Goh",
    "age": "29",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ท่ากบ 100 เมตรหญิง",
    "medal": "ทองแดง",
    "countryBorn": "Singapore"
  },
  {
    "number": 311,
    "year": 2020,
    "city": "Tokyo",
    "name": "Yip Pin Xiu",
    "age": "29",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ท่ากบ 100 เมตรหญิง",
    "medal": "ทอง",
    "countryBorn": "Singapore"
  },
  {
    "number": 312,
    "year": 2020,
    "city": "Tokyo",
    "name": "Yip Pin Xiu",
    "age": "29",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ท่ากบ 50 เมตรหญิง",
    "medal": "ทอง",
    "countryBorn": "Singapore"
  },
  {
    "number": 313,
    "year": 2024,
    "city": "Paris",
    "name": "Yip Pin Xiu",
    "age": "32",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ท่ากบ 100 เมตรหญิง",
    "medal": "ทอง",
    "countryBorn": "Singapore"
  },
  {
    "number": 314,
    "year": 2024,
    "city": "Paris",
    "name": "Yip Pin Xiu",
    "age": "32",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ท่ากบ 50 เมตรหญิง",
    "medal": "ทอง",
    "countryBorn": "Singapore"
  },
  {
    "number": 315,
    "year": 2024,
    "city": "Paris",
    "name": "Jeralyn Tan",
    "age": "35",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "บอคเซีย",
    "sportSub": "ประเภทบุคคลหญิง BC1",
    "medal": "เงิน",
    "countryBorn": "Singapore"
  },
  {
    "number": 316,
    "year": 2000,
    "city": "Sydney",
    "name": "Trần Hiếu Ngân",
    "age": "26",
    "gender": "หญิง",
    "competition": "Olympic",
    "sport": "เทควันโด",
    "sportSub": "หญิง 57 กก.",
    "medal": "เงิน",
    "countryBorn": "Vietnam"
  },
  {
    "number": 317,
    "year": 2008,
    "city": "Beijing",
    "name": "Hoàng Anh Tuấn",
    "age": "22",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย 56 กก.",
    "medal": "เงิน",
    "countryBorn": "Vietnam"
  },
  {
    "number": 318,
    "year": 2012,
    "city": "London",
    "name": "Trần Lê Quốc Toàn",
    "age": "23",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย 56 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Vietnam"
  },
  {
    "number": 319,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Hoàng Xuân Vinh",
    "age": "41",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยิงปืน",
    "sportSub": "ปืนลม ชาย ระยะ 10 เมตร",
    "medal": "ทอง",
    "countryBorn": "Vietnam"
  },
  {
    "number": 320,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Hoàng Xuân Vinh",
    "age": "41",
    "gender": "ชาย",
    "competition": "Olympic",
    "sport": "ยิงปืน",
    "sportSub": "ปืนลม ชาย ระยะ 50 เมตร",
    "medal": "เงิน",
    "countryBorn": "Vietnam"
  },
  {
    "number": 321,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Lê Văn Công",
    "age": "32",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย 49 กก.",
    "medal": "ทอง",
    "countryBorn": "Vietnam"
  },
  {
    "number": 322,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Võ Thanh Tùng",
    "age": "31",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "ว่ายน้ำ",
    "sportSub": "ฟรีสไตล์ 50 เมตร ชาย",
    "medal": "เงิน",
    "countryBorn": "Vietnam"
  },
  {
    "number": 323,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Dang Thi Linh Phuong",
    "age": "33",
    "gender": "หญิง",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "หญิง 50 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Vietnam"
  },
  {
    "number": 324,
    "year": 2016,
    "city": "Rio de Janeiro",
    "name": "Cao Ngọc Hùng",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "พุ่งเหลนชาย",
    "medal": "ทองแดง",
    "countryBorn": "Vietnam"
  },
  {
    "number": 325,
    "year": 2020,
    "city": "Tokyo",
    "name": "Lê Văn Công",
    "age": "37",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย 49 กก.",
    "medal": "เงิน",
    "countryBorn": "Vietnam"
  },
  {
    "number": 326,
    "year": 2024,
    "city": "Paris",
    "name": "Lê Văn Công",
    "age": "40",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย 49 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Vietnam"
  },
  {
    "number": 327,
    "year": 1976,
    "city": "Toronto",
    "name": "Tin Ngwe",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วิ่ง 100 เมตรชาย",
    "medal": "ทอง",
    "countryBorn": "Myanmar"
  },
  {
    "number": 328,
    "year": 1976,
    "city": "Toronto",
    "name": "Aung Than",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วิ่ง 100 เมตรชาย",
    "medal": "เงิน",
    "countryBorn": "Myanmar"
  },
  {
    "number": 329,
    "year": 1976,
    "city": "Toronto",
    "name": "Tin Win",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "วิ่ง 100 เมตรชาย",
    "medal": "ทองแดง",
    "countryBorn": "Myanmar"
  },
  {
    "number": 330,
    "year": 1984,
    "city": "New York/Stoke Mandeville",
    "name": "Tin Ngwe",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "กระโดดสูงชาย",
    "medal": "ทอง",
    "countryBorn": "Myanmar"
  },
  {
    "number": 331,
    "year": 1984,
    "city": "New York/Stoke Mandeville",
    "name": "Tin Ngwe",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "กระโดดไกลชาย",
    "medal": "เงิน",
    "countryBorn": "Myanmar"
  },
  {
    "number": 332,
    "year": 1984,
    "city": "New York/Stoke Mandeville",
    "name": "Aung Gyi",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "กระโดดไกลชาย",
    "medal": "เงิน",
    "countryBorn": "Myanmar"
  },
  {
    "number": 333,
    "year": 1984,
    "city": "New York/Stoke Mandeville",
    "name": "Aung Gyi",
    "age": "13",
    "gender": "ไม่ระบุ",
    "competition": "Paralympic",
    "sport": "กรีฑา",
    "sportSub": "กระโดดสูงชาย",
    "medal": "ทองแดง",
    "countryBorn": "Myanmar"
  },
  {
    "number": 334,
    "year": 2008,
    "city": "Beijing",
    "name": "Eay Simay",
    "age": "27",
    "gender": "ชาย",
    "competition": "Paralympic",
    "sport": "ยกน้ำหนัก",
    "sportSub": "ชาย น้ำหนัก 48 กก.",
    "medal": "ทองแดง",
    "countryBorn": "Laos"
  }
];

// ==========================================================================
// 2. ข้อมูลอ้างอิงประเทศเจ้าภาพ และประเทศเกิด (Host Countries & Flags)
// ==========================================================================
const HOST_COUNTRY_MAP = {
  'amsterdam': { city: 'Amsterdam', countryTh: 'เนเธอร์แลนด์', countryEn: 'Netherlands', code: 'nl', emoji: '🇳🇱' },
  'los angeles': { city: 'Los Angeles', countryTh: 'สหรัฐอเมริกา', countryEn: 'United States', code: 'us', emoji: '🇺🇸' },
  'berlin': { city: 'Berlin', countryTh: 'เยอรมนี', countryEn: 'Germany', code: 'de', emoji: '🇩🇪' },
  'rome': { city: 'Rome', countryTh: 'อิตาลี', countryEn: 'Italy', code: 'it', emoji: '🇮🇹' },
  'tokyo': { city: 'Tokyo', countryTh: 'ญี่ปุ่น', countryEn: 'Japan', code: 'jp', emoji: '🇯🇵' },
  'montreal': { city: 'Montreal', countryTh: 'แคนาดา', countryEn: 'Canada', code: 'ca', emoji: '🇨🇦' },
  'toronto': { city: 'Toronto', countryTh: 'แคนาดา', countryEn: 'Canada', code: 'ca', emoji: '🇨🇦' },
  'arnhem': { city: 'Arnhem', countryTh: 'เนเธอร์แลนด์', countryEn: 'Netherlands', code: 'nl', emoji: '🇳🇱' },
  'new york/stoke mandeville': { city: 'New York / Stoke Mandeville', countryTh: 'สหรัฐฯ / สหราชอาณาจักร', countryEn: 'USA / UK', code: 'us', emoji: '🇺🇸' },
  'seoul': { city: 'Seoul', countryTh: 'เกาหลีใต้', countryEn: 'South Korea', code: 'kr', emoji: '🇰🇷' },
  'barcelona': { city: 'Barcelona', countryTh: 'สเปน', countryEn: 'Spain', code: 'es', emoji: '🇪🇸' },
  'atlanta': { city: 'Atlanta', countryTh: 'สหรัฐอเมริกา', countryEn: 'United States', code: 'us', emoji: '🇺🇸' },
  'sydney': { city: 'Sydney', countryTh: 'ออสเตรเลีย', countryEn: 'Australia', code: 'au', emoji: '🇦🇺' },
  'athens': { city: 'Athens', countryTh: 'กรีซ', countryEn: 'Greece', code: 'gr', emoji: '🇬🇷' },
  'beijing': { city: 'Beijing', countryTh: 'จีน', countryEn: 'China', code: 'cn', emoji: '🇨🇳' },
  'london': { city: 'London', countryTh: 'สหราชอาณาจักร', countryEn: 'United Kingdom', code: 'gb', emoji: '🇬🇧' },
  'rio de janeiro': { city: 'Rio de Janeiro', countryTh: 'บราซิล', countryEn: 'Brazil', code: 'br', emoji: '🇧🇷' },
  'paris': { city: 'Paris', countryTh: 'ฝรั่งเศส', countryEn: 'France', code: 'fr', emoji: '🇫🇷' }
};

const BIRTH_COUNTRY_MAP = {
  'Thailand': { nameTh: 'ไทย', code: 'th', emoji: '🇹🇭' },
  'Indonesia': { nameTh: 'อินโดนีเซีย', code: 'id', emoji: '🇮🇩' },
  'Malaysia': { nameTh: 'มาเลเซีย', code: 'my', emoji: '🇲🇾' },
  'Philippines': { nameTh: 'ฟิลิปปินส์', code: 'ph', emoji: '🇵🇭' },
  'Singapore': { nameTh: 'สิงคโปร์', code: 'sg', emoji: '🇸🇬' },
  'Vietnam': { nameTh: 'เวียดนาม', code: 'vn', emoji: '🇻🇳' },
  'Myanmar': { nameTh: 'เมียนมา', code: 'mm', emoji: '🇲🇲' },
  'Laos': { nameTh: 'ลาว', code: 'la', emoji: '🇱🇦' }
};

/**
 * ดึงข้อมูลประเทศเจ้าภาพจากชื่อเมือง
 */
function getHostInfo(city) {
  const key = (city || '').toLowerCase().trim();
  if (HOST_COUNTRY_MAP[key]) {
    return HOST_COUNTRY_MAP[key];
  }
  return { city: city, countryTh: city, countryEn: city, code: '', emoji: '🏳️' };
}

/**
 * ดึงข้อมูลประเทศเกิด
 */
function getBirthCountryInfo(country) {
  const key = (country || '').trim();
  if (BIRTH_COUNTRY_MAP[key]) {
    return BIRTH_COUNTRY_MAP[key];
  }
  return { nameTh: key, code: '', emoji: '🏳️' };
}

/**
 * สร้าง HTML สำหรับแสดงธงชาติและชื่อประเทศ
 */
function renderCountryFlagBadge(code, name, subText = '') {
  if (!code) {
    return `<span>${escapeHtml(name)}</span>`;
  }
  const displayText = subText ? `${escapeHtml(name)} (${escapeHtml(subText)})` : escapeHtml(name);
  return `
    <span class="country-badge">
      <img 
        src="https://flagcdn.com/20x15/${code.toLowerCase()}.png" 
        srcset="https://flagcdn.com/40x30/${code.toLowerCase()}.png 2x" 
        width="20" 
        height="14" 
        alt="${escapeHtml(name)}" 
        class="country-flag-img"
        loading="lazy"
      >
      <span>${displayText}</span>
    </span>
  `;
}

// ==========================================================================
// 3. สถานะของแอปพลิเคชัน (Application State)
// ==========================================================================
const state = {
  filters: {
    search: '',
    year: 'all',
    competition: 'all',
    sport: 'all',
    medal: 'all',
    country: 'all'
  },
  sort: {
    column: 'year',
    direction: 'desc' // 'desc' หรือ 'asc'
  },
  pagination: {
    currentPage: 1,
    itemsPerPage: 10
  },
  charts: {
    doughnut: null,
    stackedBar: null,
    topSports: null
  }
};

// ==========================================================================
// 4. ฟังก์ชันเริ่มต้นระบบ (Initialization)
// ==========================================================================
document.addEventListener('DOMContentLoaded', () => {
  initDropdowns();
  initEventListeners();
  initCharts();
  renderApp();
  
  if (window.lucide) {
    window.lucide.createIcons();
  }
});

/**
 * เติมตัวเลือกใน Dropdown ตัวกรองปี กีฬา และประเทศเกิด
 */
function initDropdowns() {
  const yearSelect = document.getElementById('filter-year');
  const sportSelect = document.getElementById('filter-sport');
  const countrySelect = document.getElementById('filter-country');

  // 1. เติมปี พร้อมระบุประเทศเจ้าภาพ
  const uniqueYears = Array.from(new Set(OLYMPIC_DATA.map(d => d.year)))
    .sort((a, b) => b - a);

  uniqueYears.forEach(year => {
    // หาเมืองเจ้าภาพของปีนั้น
    const sample = OLYMPIC_DATA.find(d => d.year === year);
    const host = sample ? getHostInfo(sample.city) : { countryTh: '', emoji: '' };
    const opt = document.createElement('option');
    opt.value = year.toString();
    opt.textContent = `ปี ${year} - ${host.city || ''} (${host.countryTh}) ${host.emoji}`;
    yearSelect.appendChild(opt);
  });

  // 2. เติมชนิดกีฬา
  const collator = new Intl.Collator('th');
  const uniqueSports = Array.from(new Set(OLYMPIC_DATA.map(d => d.sport)))
    .sort((a, b) => collator.compare(a, b));

  uniqueSports.forEach(sport => {
    const opt = document.createElement('option');
    opt.value = sport;
    opt.textContent = sport;
    sportSelect.appendChild(opt);
  });

  // 3. เติมตัวกรองประเทศเกิด (พร้อมธงชาติ Emoji และชื่อภาษาไทย)
  const uniqueCountries = Array.from(new Set(OLYMPIC_DATA.map(d => d.countryBorn)))
    .filter(c => c && c !== 'ไม่ระบุ')
    .sort((a, b) => a.localeCompare(b));

  uniqueCountries.forEach(country => {
    const info = getBirthCountryInfo(country);
    const opt = document.createElement('option');
    opt.value = country;
    opt.textContent = `${info.emoji} ${country} (${info.nameTh})`;
    countrySelect.appendChild(opt);
  });
}

/**
 * ผูกเหตุการณ์ (Event Listeners) เข้ากับองค์ประกอบ UI
 */
function initEventListeners() {
  // ช่องค้นหาชื่อนักกีฬา
  const searchInput = document.getElementById('filter-search');
  searchInput.addEventListener('input', (e) => {
    state.filters.search = e.target.value.trim().toLowerCase();
    state.pagination.currentPage = 1;
    renderApp();
  });

  // ตัวกรองปี
  const yearSelect = document.getElementById('filter-year');
  yearSelect.addEventListener('change', (e) => {
    state.filters.year = e.target.value;
    state.pagination.currentPage = 1;
    renderApp();
  });

  // ตัวกรองการแข่งขัน
  const compSelect = document.getElementById('filter-competition');
  compSelect.addEventListener('change', (e) => {
    state.filters.competition = e.target.value;
    state.pagination.currentPage = 1;
    renderApp();
  });

  // ตัวกรองกีฬา
  const sportSelect = document.getElementById('filter-sport');
  sportSelect.addEventListener('change', (e) => {
    state.filters.sport = e.target.value;
    state.pagination.currentPage = 1;
    renderApp();
  });

  // ตัวกรองเหรียญรางวัล
  const medalSelect = document.getElementById('filter-medal');
  medalSelect.addEventListener('change', (e) => {
    state.filters.medal = e.target.value;
    state.pagination.currentPage = 1;
    renderApp();
  });

  // ตัวกรองประเทศเกิด (Country of Birth)
  const countrySelect = document.getElementById('filter-country');
  countrySelect.addEventListener('change', (e) => {
    state.filters.country = e.target.value;
    state.pagination.currentPage = 1;
    renderApp();
  });

  // ปุ่มล้างตัวกรอง
  const resetBtn = document.getElementById('btn-reset-filters');
  resetBtn.addEventListener('click', resetAllFilters);


  // ตัวเลือกจำนวนรายการต่อหน้า (Page Size)
  const pageSizeSelect = document.getElementById('select-page-size');
  pageSizeSelect.addEventListener('change', (e) => {
    state.pagination.itemsPerPage = parseInt(e.target.value, 10);
    state.pagination.currentPage = 1;
    renderApp();
  });

  // การเรียงลำดับหัวตาราง (Sortable Table Headers)
  const sortHeaders = document.querySelectorAll('.th-sortable');
  sortHeaders.forEach(th => {
    const handleSort = () => {
      const col = th.dataset.sort;
      if (state.sort.column === col) {
        state.sort.direction = state.sort.direction === 'asc' ? 'desc' : 'asc';
      } else {
        state.sort.column = col;
        state.sort.direction = col === 'year' ? 'desc' : 'asc';
      }
      renderApp();
    };

    th.addEventListener('click', handleSort);
    th.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        handleSort();
      }
    });
  });

  // ปุ่มหน้าก่อนหน้า / หน้าถัดไป
  const prevBtn = document.getElementById('btn-prev-page');
  const nextBtn = document.getElementById('btn-next-page');

  prevBtn.addEventListener('click', () => {
    if (state.pagination.currentPage > 1) {
      state.pagination.currentPage--;
      renderApp();
    }
  });

  nextBtn.addEventListener('click', () => {
    const totalFiltered = getFilteredData().length;
    const totalPages = Math.ceil(totalFiltered / state.pagination.itemsPerPage) || 1;
    if (state.pagination.currentPage < totalPages) {
      state.pagination.currentPage++;
      renderApp();
    }
  });

  // การทำงานของ Modal ปิดหน้าต่าง
  const modal = document.getElementById('athlete-modal');
  const closeModalBtn = document.getElementById('btn-close-modal');
  const dismissModalBtn = document.getElementById('btn-modal-dismiss');

  const hideModal = () => {
    modal.setAttribute('hidden', '');
    document.body.style.overflow = '';
  };

  closeModalBtn.addEventListener('click', hideModal);
  dismissModalBtn.addEventListener('click', hideModal);

  modal.addEventListener('click', (e) => {
    if (e.target === modal) {
      hideModal();
    }
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !modal.hasAttribute('hidden')) {
      hideModal();
    }
  });

  // ไฮไลต์เมนูตามการเลื่อนหน้าจอ
  window.addEventListener('scroll', updateActiveNavOnScroll);
}

/**
 * ล้างค่าตัวกรองทั้งหมดเป็นค่าเริ่มต้น
 */
function resetAllFilters() {
  state.filters.search = '';
  state.filters.year = 'all';
  state.filters.competition = 'all';
  state.filters.sport = 'all';
  state.filters.medal = 'all';
  state.filters.country = 'all';
  state.pagination.currentPage = 1;

  document.getElementById('filter-search').value = '';
  document.getElementById('filter-year').value = 'all';
  document.getElementById('filter-competition').value = 'all';
  document.getElementById('filter-sport').value = 'all';
  document.getElementById('filter-medal').value = 'all';
  document.getElementById('filter-country').value = 'all';

  renderApp();
}

// ==========================================================================
// 5. การกรองและการเรียงลำดับข้อมูล (Filtering & Sorting)
// ==========================================================================

/**
 * กรองข้อมูลตามเงื่อนไขที่ผู้ใช้เลือก
 */
function getFilteredData() {
  return OLYMPIC_DATA.filter(item => {
    // 1. ค้นหาชื่อนักกีฬา หรือชื่อทีม
    if (state.filters.search) {
      const q = state.filters.search;
      const host = getHostInfo(item.city);
      const matchName = item.name.toLowerCase().includes(q);
      const matchSport = item.sport.toLowerCase().includes(q);
      const matchSub = item.sportSub.toLowerCase().includes(q);
      const matchCity = item.city.toLowerCase().includes(q);
      const matchHostCountry = host.countryTh.toLowerCase().includes(q) || host.countryEn.toLowerCase().includes(q);
      const matchBirthCountry = item.countryBorn.toLowerCase().includes(q);
      if (!matchName && !matchSport && !matchSub && !matchCity && !matchHostCountry && !matchBirthCountry) {
        return false;
      }
    }

    // 2. กรองปี
    if (state.filters.year !== 'all') {
      if (item.year.toString() !== state.filters.year) {
        return false;
      }
    }

    // 3. กรองประเภทการแข่งขัน (Olympic / Paralympic)
    if (state.filters.competition !== 'all') {
      if (item.competition !== state.filters.competition) {
        return false;
      }
    }

    // 4. กรองชนิดกีฬา
    if (state.filters.sport !== 'all') {
      if (item.sport !== state.filters.sport) {
        return false;
      }
    }

    // 5. กรองชนิดเหรียญ (ทอง / เงิน / ทองแดง)
    if (state.filters.medal !== 'all') {
      if (item.medal !== state.filters.medal) {
        return false;
      }
    }

    // 6. กรองประเทศเกิด (Country of Birth)
    if (state.filters.country !== 'all') {
      if (item.countryBorn !== state.filters.country) {
        return false;
      }
    }

    return true;
  });
}

/**
 * เรียงลำดับข้อมูล
 */
function getSortedData(data) {
  const { column, direction } = state.sort;
  const modifier = direction === 'asc' ? 1 : -1;
  const collator = new Intl.Collator('th');

  return [...data].sort((a, b) => {
    if (column === 'year') {
      return (a.year - b.year) * modifier;
    }
    if (column === 'name') {
      return collator.compare(a.name, b.name) * modifier;
    }
    if (column === 'medal') {
      const rank = { 'ทอง': 1, 'เงิน': 2, 'ทองแดง': 3 };
      const rankA = rank[a.medal] || 99;
      const rankB = rank[b.medal] || 99;
      return (rankA - rankB) * modifier;
    }
    return 0;
  });
}

// ==========================================================================
// 6. การแสดงผล UI (Render Pipeline)
// ==========================================================================

/**
 * ฟังก์ชันหลักในการอัปเดตองค์ประกอบทั้งหมดบนหน้า
 */
function renderApp() {
  const filteredData = getFilteredData();
  const sortedData = getSortedData(filteredData);

  updateStatCards(filteredData);
  updateActiveTagsBar();
  renderTable(sortedData, filteredData.length);
  renderPagination(filteredData.length);
  updateSortHeaderUI();
  updateCharts(filteredData);

  if (window.lucide) {
    window.lucide.createIcons();
  }
}

/**
 * อัปเดตการ์ดสถิติ 4 ใบ
 */
function updateStatCards(data) {
  let gold = 0;
  let silver = 0;
  let bronze = 0;

  data.forEach(item => {
    if (item.medal === 'ทอง') gold++;
    else if (item.medal === 'เงิน') silver++;
    else if (item.medal === 'ทองแดง') bronze++;
  });

  const total = data.length;

  document.getElementById('stat-total-val').textContent = total.toLocaleString();
  document.getElementById('stat-gold-val').textContent = gold.toLocaleString();
  document.getElementById('stat-silver-val').textContent = silver.toLocaleString();
  document.getElementById('stat-bronze-val').textContent = bronze.toLocaleString();
}

/**
 * อัปเดตแถบป้ายตัวกรองที่ใช้งานอยู่
 */
function updateActiveTagsBar() {
  const container = document.getElementById('active-tags-container');
  container.innerHTML = '';

  const activeFilters = [];

  if (state.filters.search) {
    activeFilters.push({
      label: `ค้นหา: "${state.filters.search}"`,
      clear: () => {
        state.filters.search = '';
        document.getElementById('filter-search').value = '';
      }
    });
  }

  if (state.filters.year !== 'all') {
    activeFilters.push({
      label: `ปี: ${state.filters.year}`,
      clear: () => {
        state.filters.year = 'all';
        document.getElementById('filter-year').value = 'all';
      }
    });
  }

  if (state.filters.competition !== 'all') {
    activeFilters.push({
      label: `การแข่งขัน: ${state.filters.competition}`,
      clear: () => {
        state.filters.competition = 'all';
        document.getElementById('filter-competition').value = 'all';
      }
    });
  }

  if (state.filters.sport !== 'all') {
    activeFilters.push({
      label: `กีฬา: ${state.filters.sport}`,
      clear: () => {
        state.filters.sport = 'all';
        document.getElementById('filter-sport').value = 'all';
      }
    });
  }

  if (state.filters.medal !== 'all') {
    activeFilters.push({
      label: `เหรียญ: ${state.filters.medal}`,
      clear: () => {
        state.filters.medal = 'all';
        document.getElementById('filter-medal').value = 'all';
      }
    });
  }

  if (state.filters.country !== 'all') {
    const info = getBirthCountryInfo(state.filters.country);
    activeFilters.push({
      label: `ประเทศเกิด: ${info.emoji} ${state.filters.country}`,
      clear: () => {
        state.filters.country = 'all';
        document.getElementById('filter-country').value = 'all';
      }
    });
  }

  if (activeFilters.length === 0) {
    container.innerHTML = '<span class="tag-default">แสดงข้อมูลทั้งหมด</span>';
    return;
  }

  activeFilters.forEach(f => {
    const tag = document.createElement('span');
    tag.className = 'active-tag';
    tag.innerHTML = `
      <span>${f.label}</span>
      <button type="button" class="tag-remove" aria-label="ลบตัวกรองนี้">
        <i data-lucide="x" class="tag-remove-icon"></i>
      </button>
    `;
    tag.querySelector('.tag-remove').addEventListener('click', () => {
      f.clear();
      state.pagination.currentPage = 1;
      renderApp();
    });
    container.appendChild(tag);
  });
}

/**
 * แสดงตารางรายชื่อนักกีฬา
 */
function renderTable(data, totalCount) {
  const tbody = document.getElementById('athletes-table-body');
  const tableResponsive = document.querySelector('.table-responsive');
  const paginationBar = document.querySelector('.pagination-bar');
  const resultsInfo = document.getElementById('table-results-info');

  tbody.innerHTML = '';

  if (totalCount === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="10" class="td-no-data">
          ไม่พบข้อมูลที่ตรงกับเงื่อนไขการค้นหา
        </td>
      </tr>
    `;
    tableResponsive.style.display = 'block';
    paginationBar.style.display = 'none';
    resultsInfo.textContent = 'ไม่พบผลลัพธ์ที่ตรงกับตัวกรอง (0 รายการ)';
    return;
  }

  tableResponsive.style.display = 'block';
  paginationBar.style.display = 'flex';

  const { currentPage, itemsPerPage } = state.pagination;
  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = Math.min(startIndex + itemsPerPage, totalCount);
  const pageData = data.slice(startIndex, endIndex);

  resultsInfo.textContent = `แสดง ${(startIndex + 1).toLocaleString()} - ${endIndex.toLocaleString()} จากทั้งหมด ${totalCount.toLocaleString()} รายการ`;

  pageData.forEach(item => {
    const tr = document.createElement('tr');
    tr.tabIndex = 0;
    tr.setAttribute('role', 'button');
    tr.setAttribute('aria-label', `คลิกเพื่อดูรายละเอียดของ ${item.name}`);

    // กำหนด Badge การแข่งขัน
    const compBadgeClass = item.competition === 'Olympic' ? 'badge-olympic' : 'badge-paralympic';
    
    // กำหนด Badge เหรียญ
    let medalBadgeClass = 'badge-gold';
    if (item.medal === 'เงิน') medalBadgeClass = 'badge-silver';
    else if (item.medal === 'ทองแดง') medalBadgeClass = 'badge-bronze';

    // ข้อมูลประเทศเจ้าภาพ
    const host = getHostInfo(item.city);
    const hostBadgeHtml = renderCountryFlagBadge(host.code, `${host.city}, ${host.countryTh}`);

    // ข้อมูลประเทศเกิด
    const birthInfo = getBirthCountryInfo(item.countryBorn);
    const birthBadgeHtml = item.countryBorn === 'ไม่ระบุ' 
      ? '<span class="td-muted">ไม่ระบุ</span>' 
      : renderCountryFlagBadge(birthInfo.code, item.countryBorn, birthInfo.nameTh);

    tr.innerHTML = `
      <td class="td-year">${item.year}</td>
      <td class="td-athlete-name">${escapeHtml(item.name)}</td>
      <td>
        <span class="badge ${compBadgeClass}">${item.competition}</span>
      </td>
      <td>${escapeHtml(item.sport)}</td>
      <td>${escapeHtml(item.sportSub)}</td>
      <td>${item.age === 'ไม่ระบุ' ? '<span class="td-muted">ไม่ระบุ</span>' : escapeHtml(item.age)}</td>
      <td>${item.gender === 'ไม่ระบุ' ? '<span class="td-muted">ไม่ระบุ</span>' : escapeHtml(item.gender)}</td>
      <td>
        <span class="badge ${medalBadgeClass}">
          <span class="badge-indicator-dot"></span>
          ${escapeHtml(item.medal)}
        </span>
      </td>
      <td>${hostBadgeHtml}</td>
      <td>${birthBadgeHtml}</td>
    `;

    tr.addEventListener('click', () => openModal(item));
    tr.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        openModal(item);
      }
    });

    tbody.appendChild(tr);
  });
}

/**
 * อัปเดตไอคอนและคลาสที่หัวตารางสำหรับคอลัมน์ที่กำลังเรียงลำดับ
 */
function updateSortHeaderUI() {
  const { column, direction } = state.sort;
  const headers = document.querySelectorAll('.th-sortable');

  headers.forEach(th => {
    const col = th.dataset.sort;
    const icon = th.querySelector('.sort-icon');
    if (col === column) {
      th.classList.add('sorted');
      if (icon) {
        icon.setAttribute('data-lucide', direction === 'asc' ? 'arrow-up' : 'arrow-down');
      }
    } else {
      th.classList.remove('sorted');
      if (icon) {
        icon.setAttribute('data-lucide', 'arrow-up-down');
      }
    }
  });
}

/**
 * สร้างส่วนควบคุมการแบ่งหน้า (Pagination)
 */
function renderPagination(totalCount) {
  const { currentPage, itemsPerPage } = state.pagination;
  const totalPages = Math.ceil(totalCount / itemsPerPage) || 1;

  const infoEl = document.getElementById('pagination-info');
  infoEl.textContent = `หน้า ${currentPage.toLocaleString()} จาก ${totalPages.toLocaleString()} หน้า`;

  const prevBtn = document.getElementById('btn-prev-page');
  const nextBtn = document.getElementById('btn-next-page');

  prevBtn.disabled = currentPage <= 1;
  nextBtn.disabled = currentPage >= totalPages;

  const numbersContainer = document.getElementById('page-numbers-container');
  numbersContainer.innerHTML = '';

  const pageNumbers = getPaginationPages(currentPage, totalPages);

  pageNumbers.forEach(p => {
    if (p === '...') {
      const span = document.createElement('span');
      span.className = 'page-ellipsis';
      span.textContent = '...';
      numbersContainer.appendChild(span);
    } else {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = `btn-page-number ${p === currentPage ? 'active' : ''}`;
      btn.textContent = p;
      btn.setAttribute('aria-label', `ไปยังหน้า ${p}`);
      if (p === currentPage) {
        btn.setAttribute('aria-current', 'page');
      }
      btn.addEventListener('click', () => {
        state.pagination.currentPage = p;
        renderApp();
      });
      numbersContainer.appendChild(btn);
    }
  });
}

/**
 * คำนวณชุดตัวเลขหน้าสำหรับแสดงใน Pagination
 */
function getPaginationPages(current, total) {
  if (total <= 7) {
    return Array.from({ length: total }, (_, i) => i + 1);
  }

  if (current <= 3) {
    return [1, 2, 3, 4, '...', total];
  }

  if (current >= total - 2) {
    return [1, '...', total - 3, total - 2, total - 1, total];
  }

  return [1, '...', current - 1, current, current + 1, '...', total];
}

// ==========================================================================
// 7. หน้าต่างแสดงรายละเอียด (Modal)
// ==========================================================================

/**
 * เปิดหน้าต่าง Modal แสดงรายละเอียดครบทั้ง 10 คอลัมน์ และประเทศเจ้าภาพพร้อมธงชาติ
 */
function openModal(item) {
  const modal = document.getElementById('athlete-modal');

  // ข้อมูลส่วนหัว Modal
  document.getElementById('modal-athlete-name').textContent = item.name;

  const compBadge = document.getElementById('modal-competition-badge');
  compBadge.className = `badge ${item.competition === 'Olympic' ? 'badge-olympic' : 'badge-paralympic'}`;
  compBadge.textContent = item.competition;

  const medalBadge = document.getElementById('modal-medal-badge');
  let medalClass = 'badge-gold';
  let iconClass = 'icon-gold';

  if (item.medal === 'เงิน') {
    medalClass = 'badge-silver';
    iconClass = 'icon-silver';
  } else if (item.medal === 'ทองแดง') {
    medalClass = 'badge-bronze';
    iconClass = 'icon-bronze';
  }

  medalBadge.className = `badge ${medalClass}`;
  medalBadge.textContent = `เหรียญ${item.medal}`;

  const iconBox = document.getElementById('modal-medal-icon-box');
  iconBox.className = `modal-icon-badge ${iconClass}`;

  // ข้อมูลประเทศเจ้าภาพ
  const host = getHostInfo(item.city);
  const hostHtml = renderCountryFlagBadge(host.code, host.countryTh, host.countryEn);

  // ข้อมูลประเทศเกิด
  const birthInfo = getBirthCountryInfo(item.countryBorn);
  const birthHtml = item.countryBorn === 'ไม่ระบุ' 
    ? '<span class="td-muted">ไม่ระบุ</span>' 
    : renderCountryFlagBadge(birthInfo.code, item.countryBorn, birthInfo.nameTh);

  // ข้อมูลในกริด Modal
  document.getElementById('modal-year').textContent = item.year;
  document.getElementById('modal-competition').textContent = item.competition;
  document.getElementById('modal-city').textContent = host.city;
  document.getElementById('modal-host-country').innerHTML = hostHtml;
  document.getElementById('modal-medal').textContent = `เหรียญ${item.medal}`;
  document.getElementById('modal-sport').textContent = item.sport;
  document.getElementById('modal-sport-sub').textContent = item.sportSub;
  document.getElementById('modal-gender').textContent = item.gender;
  document.getElementById('modal-age').textContent = item.age;
  document.getElementById('modal-country').innerHTML = birthHtml;

  // เปิดแสดง Modal
  modal.removeAttribute('hidden');
  document.body.style.overflow = 'hidden';

  if (window.lucide) {
    window.lucide.createIcons();
  }
}

// ==========================================================================
// 8. ระบบกราฟ Chart.js (Charts)
// ==========================================================================

/**
 * เริ่มต้นสร้างกราฟทั้ง 3 กราฟ
 */
function initCharts() {
  Chart.defaults.font.family = "'Kanit', sans-serif";
  Chart.defaults.color = '#728096';

  // 1. Doughnut Chart: สัดส่วนเหรียญรางวัล
  const doughnutCtx = document.getElementById('doughnut-chart').getContext('2d');
  state.charts.doughnut = new Chart(doughnutCtx, {
    type: 'doughnut',
    data: {
      labels: ['เหรียญทอง', 'เหรียญเงิน', 'เหรียญทองแดง'],
      datasets: [{
        data: [0, 0, 0],
        backgroundColor: ['#E5AE24', '#8C9AAF', '#B8754D'],
        borderColor: '#FFFFFF',
        borderWidth: 2,
        hoverOffset: 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            padding: 16,
            usePointStyle: true,
            pointStyle: 'circle',
            font: {
              size: 13,
              family: "'Kanit', sans-serif"
            }
          }
        },
        tooltip: {
          callbacks: {
            label: function(context) {
              const label = context.label || '';
              const value = context.parsed || 0;
              const total = context.dataset.data.reduce((a, b) => a + b, 0);
              const percentage = total > 0 ? ((value / total) * 100).toFixed(1) : 0;
              return ` ${label}: ${value.toLocaleString()} รายการ (${percentage}%)`;
            }
          }
        }
      },
      cutout: '62%'
    }
  });

  // 2. Stacked Bar Chart: Olympic vs Paralympic
  const barCtx = document.getElementById('stacked-bar-chart').getContext('2d');
  state.charts.stackedBar = new Chart(barCtx, {
    type: 'bar',
    data: {
      labels: ['Olympic', 'Paralympic'],
      datasets: [
        {
          label: 'เหรียญทอง',
          data: [0, 0],
          backgroundColor: '#E5AE24',
          borderRadius: 4
        },
        {
          label: 'เหรียญเงิน',
          data: [0, 0],
          backgroundColor: '#8C9AAF',
          borderRadius: 4
        },
        {
          label: 'เหรียญทองแดง',
          data: [0, 0],
          backgroundColor: '#B8754D',
          borderRadius: 4
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: {
          stacked: true,
          grid: {
            display: false
          },
          ticks: {
            font: {
              size: 13,
              weight: '500',
              family: "'Kanit', sans-serif"
            },
            color: '#243B64'
          }
        },
        y: {
          stacked: true,
          beginAtZero: true,
          grid: {
            color: '#EDF2F7'
          },
          ticks: {
            precision: 0,
            font: {
              family: "'Kanit', sans-serif"
            }
          }
        }
      },
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            padding: 16,
            usePointStyle: true,
            pointStyle: 'circle',
            font: {
              size: 13,
              family: "'Kanit', sans-serif"
            }
          }
        },
        tooltip: {
          callbacks: {
            label: function(context) {
              const label = context.dataset.label || '';
              const value = context.parsed.y || 0;
              return ` ${label}: ${value.toLocaleString()} รายการ`;
            }
          }
        }
      }
    }
  });

  // 3. Top Sports Horizontal Stacked Bar Chart
  const topSportsCtx = document.getElementById('top-sports-chart').getContext('2d');
  state.charts.topSports = new Chart(topSportsCtx, {
    type: 'bar',
    data: {
      labels: [],
      datasets: [
        {
          label: 'เหรียญทอง',
          data: [],
          backgroundColor: '#E5AE24',
          borderRadius: 3
        },
        {
          label: 'เหรียญเงิน',
          data: [],
          backgroundColor: '#8C9AAF',
          borderRadius: 3
        },
        {
          label: 'เหรียญทองแดง',
          data: [],
          backgroundColor: '#B8754D',
          borderRadius: 3
        }
      ]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: {
          stacked: true,
          beginAtZero: true,
          grid: {
            color: '#EDF2F7'
          },
          ticks: {
            precision: 0,
            font: {
              size: 11,
              family: "'Kanit', sans-serif"
            }
          }
        },
        y: {
          stacked: true,
          grid: {
            display: false
          },
          ticks: {
            font: {
              size: 11,
              family: "'Kanit', sans-serif"
            },
            color: '#243B64'
          }
        }
      },
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            padding: 14,
            usePointStyle: true,
            pointStyle: 'circle',
            font: {
              size: 12,
              family: "'Kanit', sans-serif"
            }
          }
        },
        tooltip: {
          callbacks: {
            label: function(context) {
              const label = context.dataset.label || '';
              const value = context.parsed.x || 0;
              return ` ${label}: ${value.toLocaleString()} เหรียญ`;
            },
            title: function(context) {
              return context[0].label;
            }
          }
        }
      }
    }
  });
}

/**
 * อัปเดตข้อมูลกราฟตามตัวกรองปัจจุบัน
 */
function updateCharts(data) {
  let gold = 0;
  let silver = 0;
  let bronze = 0;

  let olyGold = 0;
  let olySilver = 0;
  let olyBronze = 0;

  let paraGold = 0;
  let paraSilver = 0;
  let paraBronze = 0;

  data.forEach(item => {
    const isOlympic = item.competition === 'Olympic';

    if (item.medal === 'ทอง') {
      gold++;
      if (isOlympic) olyGold++;
      else paraGold++;
    } else if (item.medal === 'เงิน') {
      silver++;
      if (isOlympic) olySilver++;
      else paraSilver++;
    } else if (item.medal === 'ทองแดง') {
      bronze++;
      if (isOlympic) olyBronze++;
      else paraBronze++;
    }
  });

  const total = data.length;

  // 1. อัปเดต Doughnut Chart
  if (state.charts.doughnut) {
    state.charts.doughnut.data.datasets[0].data = [gold, silver, bronze];
    state.charts.doughnut.update();
  }

  document.getElementById('summary-gold-count').textContent = gold.toLocaleString();
  document.getElementById('summary-silver-count').textContent = silver.toLocaleString();
  document.getElementById('summary-bronze-count').textContent = bronze.toLocaleString();
  document.getElementById('summary-total-count').textContent = total.toLocaleString();

  // 2. อัปเดต Stacked Bar Chart
  if (state.charts.stackedBar) {
    state.charts.stackedBar.data.datasets[0].data = [olyGold, paraGold];
    state.charts.stackedBar.data.datasets[1].data = [olySilver, paraSilver];
    state.charts.stackedBar.data.datasets[2].data = [olyBronze, paraBronze];
    state.charts.stackedBar.update();
  }

  const olyTotal = olyGold + olySilver + olyBronze;
  const paraTotal = paraGold + paraSilver + paraBronze;

  document.getElementById('comp-olympic-total').textContent = `${olyTotal.toLocaleString()} เหรียญ`;
  document.getElementById('comp-olympic-breakdown').textContent = `ทอง: ${olyGold} | เงิน: ${olySilver} | ทองแดง: ${olyBronze}`;

  document.getElementById('comp-paralympic-total').textContent = `${paraTotal.toLocaleString()} เหรียญ`;
  document.getElementById('comp-paralympic-breakdown').textContent = `ทอง: ${paraGold} | เงิน: ${paraSilver} | ทองแดง: ${paraBronze}`;

  // 3. อัปเดต Top Sports Chart
  updateTopSportsChart(data);

  // 4. อัปเดต Top Athletes Leaderboard
  renderTopAthletesLeaderboard(data);
}

// ==========================================================================
// 8b. Top Sports Chart & Athlete Leaderboard
// ==========================================================================

/**
 * คำนวณอันดับประเภทกีฬาที่ได้เหรียญมากที่สุด (Top N Sports)
 */
function getTopSports(data, limit) {
  var n = limit || 10;
  var map = {};
  data.forEach(function(item) {
    var sport = item.sport;
    if (!map[sport]) {
      map[sport] = { gold: 0, silver: 0, bronze: 0, total: 0 };
    }
    if (item.medal === 'ทอง') map[sport].gold++;
    else if (item.medal === 'เงิน') map[sport].silver++;
    else if (item.medal === 'ทองแดง') map[sport].bronze++;
    map[sport].total++;
  });

  var arr = Object.keys(map).map(function(sport) {
    return { sport: sport, gold: map[sport].gold, silver: map[sport].silver, bronze: map[sport].bronze, total: map[sport].total };
  });

  arr.sort(function(a, b) {
    if (b.total !== a.total) return b.total - a.total;
    if (b.gold !== a.gold) return b.gold - a.gold;
    return b.silver - a.silver;
  });

  return arr.slice(0, n);
}

/**
 * คำนวณทำเนียบนักกีฬาที่ได้เหรียญมากที่สุด (Top N Athletes)
 */
function getTopAthletes(data, limit) {
  var n = limit || 8;
  var map = {};
  data.forEach(function(item) {
    var key = item.name;
    if (!map[key]) {
      map[key] = { name: item.name, sport: item.sport, country: item.countryBorn, gold: 0, silver: 0, bronze: 0, total: 0 };
    }
    if (item.medal === 'ทอง') map[key].gold++;
    else if (item.medal === 'เงิน') map[key].silver++;
    else if (item.medal === 'ทองแดง') map[key].bronze++;
    map[key].total++;
    // ใช้ข้อมูลล่าสุด
    if (!map[key].country || map[key].country === 'ไม่ระบุ') {
      map[key].country = item.countryBorn;
    }
    if (!map[key].sport) {
      map[key].sport = item.sport;
    }
  });

  var arr = Object.keys(map).map(function(name) { return map[name]; });

  arr.sort(function(a, b) {
    if (b.total !== a.total) return b.total - a.total;
    if (b.gold !== a.gold) return b.gold - a.gold;
    return b.silver - a.silver;
  });

  return arr.slice(0, n);
}

/**
 * อัปเดต Top Sports Horizontal Bar Chart
 */
function updateTopSportsChart(data) {
  if (!state.charts.topSports) return;

  var topSports = getTopSports(data, 10);
  var labels = topSports.map(function(s) { return s.sport; });
  var goldData = topSports.map(function(s) { return s.gold; });
  var silverData = topSports.map(function(s) { return s.silver; });
  var bronzeData = topSports.map(function(s) { return s.bronze; });

  state.charts.topSports.data.labels = labels;
  state.charts.topSports.data.datasets[0].data = goldData;
  state.charts.topSports.data.datasets[1].data = silverData;
  state.charts.topSports.data.datasets[2].data = bronzeData;
  state.charts.topSports.update();
}

/**
 * สร้างทำเนียบนักกีฬาที่ได้เหรียญมากที่สุด (Leaderboard)
 * คลิกชื่อนักกีฬาเพื่อกรองตาราง
 */
function renderTopAthletesLeaderboard(data) {
  var container = document.getElementById('top-athletes-list');
  if (!container) return;

  var topAthletes = getTopAthletes(data, 8);

  if (topAthletes.length === 0) {
    container.innerHTML = '<p style="color:var(--color-text-muted);font-size:0.85rem;text-align:center;padding:24px 0;">ไม่มีข้อมูลในชุดกรองนี้</p>';
    return;
  }

  container.innerHTML = '';
  topAthletes.forEach(function(athlete, idx) {
    var rank = idx + 1;
    var rankClass = rank === 1 ? 'rank-1' : rank === 2 ? 'rank-2' : rank === 3 ? 'rank-3' : 'rank-other';

    var birthInfo = getBirthCountryInfo(athlete.country);
    var flagHtml = '';
    if (birthInfo.code) {
      flagHtml = '<img src="https://flagcdn.com/20x15/' + birthInfo.code.toLowerCase() + '.png" width="16" height="12" alt="' + escapeHtml(athlete.country) + '" style="border-radius:2px;vertical-align:middle;margin-right:3px;" loading="lazy">';
    } else {
      flagHtml = '<span style="margin-right:3px;">' + (birthInfo.emoji || '') + '</span>';
    }

    var item = document.createElement('div');
    item.className = 'leaderboard-item';
    item.title = 'คลิกเพื่อดูรายการเหรียญทั้งหมดของ ' + athlete.name;
    item.innerHTML =
      '<div class="rank-badge ' + rankClass + '">' + rank + '</div>' +
      '<div class="leaderboard-info">' +
        '<div class="leaderboard-name">' + escapeHtml(athlete.name) + '</div>' +
        '<div class="leaderboard-meta">' + flagHtml + escapeHtml(athlete.sport) + '</div>' +
      '</div>' +
      '<div class="medal-dots">' +
        '<div class="medal-dot-group"><span class="medal-dot-label">ทอง</span><span class="medal-dot-value dot-gold">' + athlete.gold + '</span></div>' +
        '<div class="medal-dot-group"><span class="medal-dot-label">เงิน</span><span class="medal-dot-value dot-silver">' + athlete.silver + '</span></div>' +
        '<div class="medal-dot-group"><span class="medal-dot-label">แดง</span><span class="medal-dot-value dot-bronze">' + athlete.bronze + '</span></div>' +
      '</div>' +
      '<span class="leaderboard-total">' + athlete.total + ' เหรียญ</span>';

    // คลิกชื่อนักกีฬา → กรองตารางตามชื่อ
    item.addEventListener('click', function() {
      var searchInput = document.getElementById('filter-search');
      state.filters.search = athlete.name.toLowerCase();
      searchInput.value = athlete.name;
      state.pagination.currentPage = 1;
      renderApp();
      // เลื่อนไปยังตาราง
      var tableSection = document.getElementById('athletes-section');
      if (tableSection) {
        tableSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    });

    container.appendChild(item);
  });
}

// ==========================================================================
// 9. การนำทางและฟังก์ชันช่วยเหลือ (Navigation & Helpers)
// ==========================================================================

/**
 * ตรวจสอบตำแหน่ง Scroll และปรับคลาส active ของเมนูนำทาง
 */
function updateActiveNavOnScroll() {
  const sections = [
    { id: 'overview', link: document.querySelector('.nav-link[href="#overview"]') },
    { id: 'athletes-section', link: document.querySelector('.nav-link[href="#athletes-section"]') },
    { id: 'charts-section', link: document.querySelector('.nav-link[href="#charts-section"]') }
  ];

  const scrollY = window.pageYOffset || document.documentElement.scrollTop;

  sections.forEach(sec => {
    const el = document.getElementById(sec.id);
    if (!el || !sec.link) return;

    const top = el.offsetTop - 120;
    const height = el.offsetHeight;

    if (scrollY >= top && scrollY < top + height) {
      document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
      sec.link.classList.add('active');
    }
  });
}

/**
 * หลีกเลี่ยง XSS โดยแปลงอักขระพิเศษเป็น HTML entities
 */
function escapeHtml(str) {
  if (typeof str !== 'string') return str;
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
