/**
 * Competency Assessment 5.0 Dashboard - Light Mode Application Controller
 * Handles PPTX Summary Analytics, JD Competency Matrix (OJT List), Interactive Charts & IDP Management
 */

// --------------------------------------------------------------------------
// 1. DATA STATE & BACKEND CONTRACT (Extracted from PPTX & JD List XLSX)
// --------------------------------------------------------------------------
const CompetencyData = {
  metadata: {
    year: 2024,
    company: "RX Company",
    totalEmployees: 281,
    corporateTargetGap: 30.0,
    groups: {
      office: { name: "เธเธฅเธธเนเธกเธเธฒเธเธชเธณเธเธฑเธเธเธฒเธ", count: 124 },
      sales: { name: "เธเธฅเธธเนเธกเธเธฒเธเธเธฒเธข", count: 102 },
      lasik: { name: "เธเธฅเธธเนเธกเธเธฒเธเน€เธฅเธชเธดเธ", count: 56 }
    },
    levels: {
      staff: { level: 1, name: "เธฃเธฐเธ”เธฑเธเธเธเธฑเธเธเธฒเธ", expectedLevel: 3, count: 233 },
      supervisor: { level: 2, name: "เธฃเธฐเธ”เธฑเธเธซเธฑเธงเธซเธเนเธฒเธเธฒเธ", expectedLevel: 4, count: 34 },
      manager: { level: 3, name: "เธฃเธฐเธ”เธฑเธเธเธนเนเธเธฑเธ”เธเธฒเธฃ", expectedLevel: 5, count: 13 }
    }
  },

  entrepreneurshipGoal: {
    target: 30.0,
    year2023: 22.67,
    year2024: 20.64,
    status: "PASSED"
  },

  coreCompetency: {
    overall: [
      { id: "relationship", name: "Relationship Excellence", total: 281, gapCount: 18, gapPct: 6.41 },
      { id: "fact", name: "Fact", total: 281, gapCount: 36, gapPct: 12.81 },
      { id: "innovative", name: "Innovative Thinking", total: 281, gapCount: 72, gapPct: 25.62 },
      { id: "learning", name: "Learning", total: 281, gapCount: 88, gapPct: 31.32 },
      { id: "entrepreneurship", name: "Entrepreneurship", total: 281, gapCount: 58, gapPct: 20.64 }
    ],
    byLevel: {
      staff: [
        { name: "Relationship", count: 233, gapCount: 3, gapPct: 1.29 },
        { name: "Fact", count: 233, gapCount: 11, gapPct: 4.72 },
        { name: "Innovative", count: 233, gapCount: 44, gapPct: 18.88 },
        { name: "Learning", count: 233, gapCount: 68, gapPct: 29.18 },
        { name: "Entrepreneurship", count: 233, gapCount: 35, gapPct: 15.02 }
      ],
      supervisor: [
        { name: "Relationship", count: 34, gapCount: 8, gapPct: 23.53 },
        { name: "Fact", count: 34, gapCount: 15, gapPct: 44.12 },
        { name: "Innovative", count: 34, gapCount: 20, gapPct: 58.82 },
        { name: "Learning", count: 34, gapCount: 13, gapPct: 38.24 },
        { name: "Entrepreneurship", count: 34, gapCount: 17, gapPct: 50.00 }
      ],
      manager: [
        { name: "Relationship", count: 14, gapCount: 7, gapPct: 50.00 },
        { name: "Fact", count: 14, gapCount: 10, gapPct: 71.43 },
        { name: "Innovative", count: 14, gapCount: 8, gapPct: 57.14 },
        { name: "Learning", count: 14, gapCount: 7, gapPct: 50.00 },
        { name: "Entrepreneurship", count: 14, gapCount: 6, gapPct: 42.86 }
      ]
    }
  },

  jobCompetency: {
    overall: [
      { name: "เธเธงเธฒเธกเธฃเธนเนเธ”เนเธฒเธเธฃเธฐเธเธเธเธธเธ“เธ เธฒเธ", total: 121, gapCount: 55, gapPct: 45.45, category: "Knowledge" },
      { name: "เธเธฒเธฃเธงเธดเน€เธเธฃเธฒเธฐเธซเนเธเนเธญเธกเธนเธฅ", total: 128, gapCount: 40, gapPct: 31.25, category: "Skill" },
      { name: "เธเธฒเธฃเธชเธทเนเธญเธชเธฒเธฃ (Communication)", total: 162, gapCount: 33, gapPct: 20.37, category: "Skill" },
      { name: "เธเธฒเธฃเนเธเนเนเธเธเธฑเธเธซเธฒเนเธฅเธฐเธเธฒเธฃเธ•เธฑเธ”เธชเธดเธเนเธ", total: 238, gapCount: 41, gapPct: 17.23, category: "Skill" },
      { name: "เธเธฒเธฃเธเธเธดเธเธฑเธ•เธดเธเธฒเธเธ•เธฒเธกเธเธฑเนเธเธ•เธญเธ & เน€เธญเธเธชเธฒเธฃ", total: 158, gapCount: 14, gapPct: 8.86, category: "Skill" }
    ]
  },

  idpActions: [
    { id: "IDP-001", name: "เธชเธกเธเธฒเธข เธกเธตเธชเธธเธ", level: "เธซเธฑเธงเธซเธเนเธฒเธเธฒเธ", group: "เธชเธณเธเธฑเธเธเธฒเธ", gapItem: "เธเธฒเธฃเธงเธดเน€เธเธฃเธฒเธฐเธซเนเธเนเธญเธกเธนเธฅ (64.71%)", status: "เธเธณเธฅเธฑเธเธ”เธณเน€เธเธดเธเธเธฒเธฃ", dueDate: "2024-12-15" },
    { id: "IDP-002", name: "เธงเธดเธ เธฒเธ”เธฒ เน€เธ”เนเธเธ”เธต", level: "เธเธนเนเธเธฑเธ”เธเธฒเธฃ", group: "เธเธฒเธข", gapItem: "เธ เธฒเธงเธฐเธเธนเนเธเธณ (81.82%)", status: "เธฃเธญเธ”เธณเน€เธเธดเธเธเธฒเธฃ", dueDate: "2024-12-30" },
    { id: "IDP-003", name: "เธญเธเธฑเธเธ•เน เธเธฑเธขเธเธเธฐ", level: "เธเธเธฑเธเธเธฒเธ", group: "เน€เธฅเธชเธดเธ", gapItem: "Learning (28.57%)", status: "เธญเธเธธเธกเธฑเธ•เธดเนเธฅเนเธง", dueDate: "2024-11-30" },
    { id: "IDP-004", name: "เธเธ เธฒเธฅเธฑเธข เธชเธงเนเธฒเธ", level: "เธเธเธฑเธเธเธฒเธ", group: "เธชเธณเธเธฑเธเธเธฒเธ", gapItem: "เธเธงเธฒเธกเธฃเธนเนเธ”เนเธฒเธเธฃเธฐเธเธเธเธธเธ“เธ เธฒเธ (46.6%)", status: "เธเธณเธฅเธฑเธเธ”เธณเน€เธเธดเธเธเธฒเธฃ", dueDate: "2024-12-20" }
  ]
};

// --------------------------------------------------------------------------
// 2. JD MATRIX DATA (Parsed from JD List _ Competency _ OJT.xlsx - 124 Positions)
// --------------------------------------------------------------------------
const JDPositionsData = [
    {
        "jdCode":  "SD-HRD-01/ACC-01",
        "dept":  "ACC",
        "title":  "ผจก.บัญชีและการเงิน",
        "group":  "Office",
        "level":  3,
        "totalCount":  13,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การสื่อสาร (Communication)",
                             "การบริหารการเงิน(Finance Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/ACC-03",
        "dept":  "ACC",
        "title":  "พนง.บัญชี",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การจัดทำบัญชี(Accounting)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/ACC-04",
        "dept":  "ACC",
        "title":  "พนง.บัญชีลูกหนี้",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การจัดทำบัญชี(Accounting)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CHD-01",
        "dept":  "CHD",
        "title":  "หน.งานแคชเชียร์",
        "group":  "Office",
        "level":  2,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การทำธุรกรรมทางธนาคาร(Banking)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CMP-01",
        "dept":  "CMP",
        "title":  "หน.งานบริหารค่าตอบแทน",
        "group":  "Office",
        "level":  2,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CMP-02",
        "dept":  "CMP",
        "title":  "พนง.Payroll",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางกฎหมาย(Legal Knowledge) ",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CRD-01",
        "dept":  "CRD",
        "title":  "ผจก.บริหารลูกหนี้",
        "group":  "Office",
        "level":  3,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CRD-02",
        "dept":  "CRD",
        "title":  "พนง. ธุรการบริหารลูกหนี้",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CRD-04",
        "dept":  "CRD",
        "title":  "พนง. เก็บเงิน กทม.",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CRD-05",
        "dept":  "CRD",
        "title":  "พนง. เก็บเงิน ตจว.",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/HRD-01",
        "dept":  "HRD",
        "title":  "ผจก. ทรัพยากรบุคคล",
        "group":  "Office",
        "level":  3,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางกฎหมาย(Legal Knowledge) ",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/HRD-04",
        "dept":  "HRD",
        "title":  "พนง. ซ่อมบำรุง",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/HRD-05",
        "dept":  "HRD",
        "title":  "พนง. ธุรการ",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/HRD-06",
        "dept":  "HRD",
        "title":  "พนง.รับโทรศัพท์",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/HRD-16",
        "dept":  "HRD",
        "title":  "พนง.ทรัพยากรบุคคล",
        "group":  "Office",
        "level":  1,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางกฎหมาย(Legal Knowledge) ",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การสรรหาบุคลากร (Recruitment)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/HRD-17",
        "dept":  "HRD",
        "title":  "พนักงานระบบคุณภาพ",
        "group":  "Office",
        "level":  1,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/HRD-18",
        "dept":  "HRD",
        "title":  "พนักงานฝึกอบรม",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/HRD-19",
        "dept":  "HRD",
        "title":  "หัวหน้างานทรัพยากรบุคคล",
        "group":  "Office",
        "level":  2,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางกฎหมาย(Legal Knowledge) ",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/HRD-20",
        "dept":  "HRD",
        "title":  "HRIS",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/ITD-04",
        "dept":  "ITD",
        "title":  "Programmer",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การแนะนำการใช้งานและซ่อมเครื่องมือแพทย์\n(Instruction and Repair medical instruments)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/ITD-05",
        "dept":  "ITD",
        "title":  "IT Support",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การแนะนำการใช้งานและซ่อมเครื่องมือแพทย์\n(Instruction and Repair medical instruments)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/ITD-09",
        "dept":  "ITD",
        "title":  "ผู้จัดการสารสนเทศ",
        "group":  "Office",
        "level":  3,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การแนะนำการใช้งานและซ่อมเครื่องมือแพทย์\n(Instruction and Repair medical instruments)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/ITD-10",
        "dept":  "ITD",
        "title":  "หัวหน้างานสารสนเทศ RX Site",
        "group":  "Office",
        "level":  2,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การแนะนำการใช้งานและซ่อมเครื่องมือแพทย์\n(Instruction and Repair medical instruments)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/ITD-11",
        "dept":  "ITD",
        "title":  "หัวหน้างานสารสนเทศ RXM Site",
        "group":  "Office",
        "level":  2,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การแนะนำการใช้งานและซ่อมเครื่องมือแพทย์\n(Instruction and Repair medical instruments)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/ITD-12",
        "dept":  "ITD",
        "title":  "หัวหน้างานสารสนเทศด้าน Infrastructure",
        "group":  "Office",
        "level":  2,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การแนะนำการใช้งานและซ่อมเครื่องมือแพทย์\n(Instruction and Repair medical instruments)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/ITD-13",
        "dept":  "ITD",
        "title":  "Programmer",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การแนะนำการใช้งานและซ่อมเครื่องมือแพทย์\n(Instruction and Repair medical instruments)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MNG-14",
        "dept":  "MNG",
        "title":  "Strategic Operation Director",
        "group":  "Office",
        "level":  5,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MNG-15",
        "dept":  "MNG",
        "title":  "Data Analyst",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MNG-16",
        "dept":  "MNG",
        "title":  "Programmer",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การใช้โปรแกรมคอมพิวเตอร์(Software\u0026Application skill)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/PSA-01",
        "dept":  "PSA",
        "title":  "หน.งานธุรการขาย",
        "group":  "Office",
        "level":  2,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/PSA-02",
        "dept":  "PSA",
        "title":  "พนง.ธุรการขาย",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/REG-03",
        "dept":  "REG",
        "title":  "Regulatory Affairs Specialist",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "การขออนุญาตเกี่ยวกับผลิตภัณฑ์ (Product Registration)",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/REG-04",
        "dept":  "REG",
        "title":  "Regulatory Affairs Supervisor",
        "group":  "Office",
        "level":  2,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "การขออนุญาตเกี่ยวกับผลิตภัณฑ์ (Product Registration)",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/ENG-02",
        "dept":  "ENG",
        "title":  "พนง.ช่าง",
        "group":  "Office",
        "level":  1,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)",
                             "การนำเสนอ(Presentation)",
                             "การแนะนำการใช้งานและซ่อมเครื่องมือแพทย์\n(Instruction and Repair medical instruments)",
                             "การสอบเทียบ(Calibration)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/ENG-03",
        "dept":  "ENG",
        "title":  "พนง.ธุรการช่าง",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การ Service รถพยาบาล"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/ENG-08",
        "dept":  "ENG",
        "title":  "Service Engineer Supervisor",
        "group":  "Office",
        "level":  2,
        "totalCount":  15,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "การแนะนำการใช้งานและซ่อมเครื่องมือแพทย์\n(Instruction and Repair medical instruments)",
                             "การสอบเทียบ(Calibration)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/REX-01",
        "dept":  "REX",
        "title":  "Ambulance \u0026 Service Manager",
        "group":  "Office",
        "level":  3,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การประกอบรถพยาบาล"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/REX-02",
        "dept":  "REX",
        "title":  "หัวหน้างานศูนย์ประกอบรถพยาบาล",
        "group":  "Office",
        "level":  2,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การประกอบรถพยาบาล"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/REX-03",
        "dept":  "REX",
        "title":  "พนง.ประกอบรถพยาบาล",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การประกอบรถพยาบาล"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/REX-04",
        "dept":  "REX",
        "title":  "พนง.ธุรการวางแผนและต้นทุน",
        "group":  "Office",
        "level":  1,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การประกอบรถพยาบาล"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/REX-05",
        "dept":  "REX",
        "title":  "พนง.ควบคุมคุณภาพรถพยาบาล",
        "group":  "Office",
        "level":  1,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การประกอบรถพยาบาล"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/REX-08",
        "dept":  "REX",
        "title":  "Product \u0026 Design Ambulance Specialist",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "การประกอบรถพยาบาล"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/PCD-06",
        "dept":  "PCD",
        "title":  "หัวหน้างานจัดซื้อ",
        "group":  "Office",
        "level":  2,
        "totalCount":  15,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การเจรจาต่อรอง(Negotiation)",
                             "การจัดซื้อ(Purchasing)",
                             "การนำเข้า-ส่งออก(Import-Export)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/PCD-09",
        "dept":  "PCD",
        "title":  "พนักงานจัดซื้อ",
        "group":  "Office",
        "level":  1,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การเจรจาต่อรอง(Negotiation)",
                             "การจัดซื้อ(Purchasing)",
                             "การนำเข้า-ส่งออก(Import-Export)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/PCD-10",
        "dept":  "PCD",
        "title":  "ผู้จัดการจัดซื้อ",
        "group":  "Office",
        "level":  3,
        "totalCount":  14,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การเจรจาต่อรอง(Negotiation)",
                             "การจัดซื้อ(Purchasing)",
                             "การนำเข้า-ส่งออก(Import-Export)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/WHD-02",
        "dept":  "WHD",
        "title":  "หน. คลังสินค้า",
        "group":  "Office",
        "level":  2,
        "totalCount":  13,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การจัดการโลจิสติกส์และซัพพลายเชนLOGISTICS-SUPPLY CHAIN MANAGEMENT",
                             "งานคลังสินค้า"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/WHD-03",
        "dept":  "WHD",
        "title":  "พนง. คลังสินค้า",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "งานคลังสินค้า"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/WHD-05",
        "dept":  "WHD",
        "title":  "พนง. ส่งสินค้า",
        "group":  "Office",
        "level":  1,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)",
                             "งานจัดส่งสินค้า"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/WHD-08",
        "dept":  "WHD",
        "title":  "พนง. ธุรการคลังสินค้า",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "งานคลังสินค้า"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/WHD-12",
        "dept":  "WHD",
        "title":  "พนง. ซ่อมบำรุง",
        "group":  "Office",
        "level":  1,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "งานคลังสินค้า",
                             "การสอบเทียบ(Calibration)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/WHD-13",
        "dept":  "WHD",
        "title":  "พนง. ธุรการจัดส่ง",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "งานจัดส่งสินค้า"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/WHD-14",
        "dept":  "WHD",
        "title":  "หัวหน้างานรับสินค้า",
        "group":  "Office",
        "level":  2,
        "totalCount":  13,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การจัดการโลจิสติกส์และซัพพลายเชนLOGISTICS-SUPPLY CHAIN MANAGEMENT",
                             "งานคลังสินค้า"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/WHD-20",
        "dept":  "WHD",
        "title":  "พนักงานรับสินค้า",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "งานคลังสินค้า"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/WHD-24",
        "dept":  "WHD",
        "title":  "หน.งานจัดส่งสินค้า",
        "group":  "Office",
        "level":  2,
        "totalCount":  13,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การจัดการโลจิสติกส์และซัพพลายเชนLOGISTICS-SUPPLY CHAIN MANAGEMENT",
                             "งานจัดส่งสินค้า"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/WHD-25",
        "dept":  "WHD",
        "title":  "ผู้จัดการคลังสินค้าและจัดส่ง",
        "group":  "Office",
        "level":  3,
        "totalCount":  13,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การจัดการโลจิสติกส์และซัพพลายเชนLOGISTICS-SUPPLY CHAIN MANAGEMENT",
                             "งานคลังสินค้า"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/WHD-27",
        "dept":  "WHD",
        "title":  "หัวหน้างานซ่อมบำรุง",
        "group":  "Office",
        "level":  2,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การสอบเทียบ(Calibration)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/Billing-01",
        "dept":  "Billing",
        "title":  "พนักงาน Billing",
        "group":  "Office",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MD-02",
        "dept":  "MD",
        "title":  "Marketing \u0026 Sales Director - Medical Device",
        "group":  "MD",
        "level":  5,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MD-03",
        "dept":  "MD",
        "title":  "Marketing Executive - MD",
        "group":  "MD",
        "level":  1,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/OPD-04",
        "dept":  "OPD",
        "title":  "Sales \u0026 Marketing Manager- OPD",
        "group":  "MD",
        "level":  4,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/SPL-04",
        "dept":  "SPL",
        "title":  "หน.งานขาย กทม.",
        "group":  "MD",
        "level":  2,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/SPL-05",
        "dept":  "SPL",
        "title":  "หน.งานขาย ตจว.",
        "group":  "MD",
        "level":  2,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/SPL-06",
        "dept":  "SPL",
        "title":  "พนง. ขาย กทม.",
        "group":  "MD",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "ทักษะการขาย(Selling Skill)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/SPL-07",
        "dept":  "SPL",
        "title":  "พนง. ขาย ตจว.",
        "group":  "MD",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "ทักษะการขาย(Selling Skill)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/SPL-10",
        "dept":  "SPL",
        "title":  "Product  Manager - SPL",
        "group":  "MD",
        "level":  3,
        "totalCount":  13,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/SPL-12",
        "dept":  "SPL",
        "title":  "Product Specialist-SPL",
        "group":  "MD",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/EED-03",
        "dept":  "EED",
        "title":  "พนง. ขาย กทม.",
        "group":  "MD",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "ทักษะการขาย(Selling Skill)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/EED-04",
        "dept":  "EED",
        "title":  "พนง. ขาย ตจว.",
        "group":  "MD",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "ทักษะการขาย(Selling Skill)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/EED-08",
        "dept":  "EED",
        "title":  "หน.งานขาย กทม.",
        "group":  "MD",
        "level":  2,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/EED-09",
        "dept":  "EED",
        "title":  "หน.งานขาย ตจว.",
        "group":  "MD",
        "level":  2,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/EED-11",
        "dept":  "EED",
        "title":  "Product  Specialist",
        "group":  "MD",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/EED-20",
        "dept":  "EED",
        "title":  "Product  Manager - EED",
        "group":  "MD",
        "level":  3,
        "totalCount":  13,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/Pharma-02",
        "dept":  "Pharma",
        "title":  "Marketing \u0026 Sales Director - Pharma",
        "group":  "Pharma",
        "level":  5,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MCD-03",
        "dept":  "MCD",
        "title":  "UPC Sales  Manager -MC",
        "group":  "Pharma",
        "level":  3,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MCD-04",
        "dept":  "MCD",
        "title":  "หน. งานขาย กทม.",
        "group":  "Pharma",
        "level":  2,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MCD-05",
        "dept":  "MCD",
        "title":  "หน. งานขาย ตจว.",
        "group":  "Pharma",
        "level":  2,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MCD-07",
        "dept":  "MCD",
        "title":  "พนง. ขาย กทม.",
        "group":  "Pharma",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การสื่อสาร (Communication)",
                             "ทักษะการขาย(Selling Skill)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MCD-08",
        "dept":  "MCD",
        "title":  "พนง. ขาย ตจว.",
        "group":  "Pharma",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การสื่อสาร (Communication)",
                             "ทักษะการขาย(Selling Skill)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MCD-13",
        "dept":  "MCD",
        "title":  "BKK Sale Manager - MC",
        "group":  "Pharma",
        "level":  3,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MCD-14",
        "dept":  "MCD",
        "title":  "Medical Product  Manager - MC",
        "group":  "Pharma",
        "level":  3,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MCD-16",
        "dept":  "MCD",
        "title":  "Medical Product Specialist-MC",
        "group":  "Pharma",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MCD-18",
        "dept":  "MCD",
        "title":  "National Sales Manager - MC",
        "group":  "Pharma",
        "level":  3,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/MCD-19",
        "dept":  "MCD",
        "title":  "Senior Sales Supervisor - MC",
        "group":  "Pharma",
        "level":  2,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-02",
        "dept":  "CSH",
        "title":  "หน. งานขาย กทม.",
        "group":  "Pharma",
        "level":  2,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-03",
        "dept":  "CSH",
        "title":  "หน. งานขาย ตจว.",
        "group":  "Pharma",
        "level":  2,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-04",
        "dept":  "CSH",
        "title":  "พนง. ขาย กทม.",
        "group":  "Pharma",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การสื่อสาร (Communication)",
                             "ทักษะการขาย(Selling Skill)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-05",
        "dept":  "CSH",
        "title":  "พนง. ขาย ตจว.",
        "group":  "Pharma",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การสื่อสาร (Communication)",
                             "ทักษะการขาย(Selling Skill)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-12",
        "dept":  "GD",
        "title":  "Graphic Designer",
        "group":  "Office",
        "level":  1,
        "totalCount":  9,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การสื่อสาร (Communication)",
                             "การออกแบบและการจัดทำสื่อ (Graphic Design and Print media Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-18",
        "dept":  "CSH",
        "title":  "Marketing Executive - CSH",
        "group":  "MKT",
        "level":  1,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-20",
        "dept":  "CSH",
        "title":  "Medical Product Specialist - CSH",
        "group":  "Pharma",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-22",
        "dept":  "CSH",
        "title":  "National Sales Manager - CSH",
        "group":  "Pharma",
        "level":  3,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-23",
        "dept":  "CSH",
        "title":  "Key Account Executive - CSH",
        "group":  "Pharma",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การสื่อสาร (Communication)",
                             "ทักษะการขาย(Selling Skill)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-26",
        "dept":  "CSH",
        "title":  "Key Account Manager - CSH",
        "group":  "Pharma",
        "level":  2,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Pharma (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-27",
        "dept":  "CSH",
        "title":  "Brand Supervisor - CSH",
        "group":  "MKT",
        "level":  2,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-28",
        "dept":  "GD",
        "title":  "Graphic Supervisor - CSH",
        "group":  "Office",
        "level":  2,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การออกแบบและการจัดทำสื่อ (Graphic Design and Print media Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-29",
        "dept":  "GD",
        "title":  "Senior Graphic Designer",
        "group":  "Office",
        "level":  1,
        "totalCount":  9,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การสื่อสาร (Communication)",
                             "การออกแบบและการจัดทำสื่อ (Graphic Design and Print media Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-31",
        "dept":  "CSH",
        "title":  "Marketing Manager - CSH",
        "group":  "MKT",
        "level":  3,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/CSH-32",
        "dept":  "CSH",
        "title":  "Content Creator - CSH",
        "group":  "MKT",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)",
                             "การออกแบบและการจัดทำสื่อ (Graphic Design and Print media Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LS-01",
        "dept":  "LS",
        "title":  "Sales \u0026 Marketing Manager - LS",
        "group":  "MD",
        "level":  4,
        "totalCount":  13,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การบริหารงานขาย (Sales Management)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LS-03",
        "dept":  "LS",
        "title":  "หน.งานขาย กทม. - LS",
        "group":  "MD",
        "level":  2,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การบริหารงานขาย (Sales Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LS-04",
        "dept":  "LS",
        "title":  "หน.งานขาย ตจว. - LS",
        "group":  "MD",
        "level":  2,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การบริหารงานขาย (Sales Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LS-05",
        "dept":  "LS",
        "title":  "พนักงานขาย กทม. - LS",
        "group":  "MD",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "ทักษะการขาย(Selling Skill)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LS-06",
        "dept":  "LS",
        "title":  "พนักงานขาย ตจว. - LS",
        "group":  "MD",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "ทักษะการขาย(Selling Skill)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LS-07",
        "dept":  "LS",
        "title":  "Product Manager - LS",
        "group":  "MD",
        "level":  3,
        "totalCount":  13,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LS-08",
        "dept":  "LS",
        "title":  "Product Speciailst - LS",
        "group":  "MD",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/GED-04",
        "dept":  "GED",
        "title":  "หน.งานขาย กทม. - GED",
        "group":  "MD",
        "level":  2,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การบริหารงานขาย (Sales Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/GED-05",
        "dept":  "GED",
        "title":  "หน.งานขาย ตจว. - GED",
        "group":  "MD",
        "level":  2,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การบริหารงานขาย (Sales Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/GED-06",
        "dept":  "GED",
        "title":  "พนักงานขาย กทม. - GED",
        "group":  "MD",
        "level":  1,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "ทักษะการขาย(Selling Skill)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/GED-07",
        "dept":  "GED",
        "title":  "พนักงานขาย ตจว. - GED",
        "group":  "MD",
        "level":  1,
        "totalCount":  11,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "ทักษะการขาย(Selling Skill)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/GED-08",
        "dept":  "GED",
        "title":  "Product Specialist - GED",
        "group":  "MD",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/GED-14",
        "dept":  "GED",
        "title":  "Sales \u0026 Marketing Manager - GED",
        "group":  "MD",
        "level":  4,
        "totalCount":  13,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การบริหารงานขาย (Sales Management)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/GED-16",
        "dept":  "GED",
        "title":  "Product Manager - GED",
        "group":  "MD",
        "level":  3,
        "totalCount":  13,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ในสินค้าและบริการ Medical  Devices (Product Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LSK-02",
        "dept":  "LSK",
        "title":  "หน.งานศูนย์เลสิก",
        "group":  "LSK",
        "level":  2,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางเลสิค(Lasik Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารจัดการ (Management)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LSK-03",
        "dept":  "LSK",
        "title":  "พยาบาลวิชาชีพประจำศูนย์เลสิก",
        "group":  "LSK",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางเลสิค(Lasik Knowledge)",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)",
                             "การพยาบาล (Nursing Practice)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LSK-04",
        "dept":  "LSK",
        "title":  "พนง.ตรวจวัดสายตา",
        "group":  "LSK",
        "level":  1,
        "totalCount":  9,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางเลสิค(Lasik Knowledge)",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LSK-05",
        "dept":  "LSK",
        "title":  "พนง.คุมเครื่องเลสิก",
        "group":  "LSK",
        "level":  1,
        "totalCount":  9,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางเลสิค(Lasik Knowledge)",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LSK-07",
        "dept":  "LSK",
        "title":  "ผู้ช่วยพยาบาล",
        "group":  "LSK",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางเลสิค(Lasik Knowledge)",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)",
                             "การพยาบาล (Nursing Practice)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LSK-11",
        "dept":  "LSK",
        "title":  "Lasik Manager",
        "group":  "LSK",
        "level":  3,
        "totalCount":  12,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางเลสิค(Lasik Knowledge)",
                             "การปฏิบัติงานเชิงกลยุทธ์(Strategic Capability )",
                             "การบริหารบุคลากร(Managing People)",
                             "การสอนงานและพัฒนาพนักงาน(Coaching)",
                             "การวิเคราะห์ข้อมูล  Data Analysis",
                             "ภาวะผู้นำ(Leadership )",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LSK-12",
        "dept":  "LSK",
        "title":  "เลขานุการแพทย์ประจำศูนย์เลสิก",
        "group":  "LSK",
        "level":  1,
        "totalCount":  9,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางเลสิค(Lasik Knowledge)",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LSK-16",
        "dept":  "LSK",
        "title":  "Lasik Consultant",
        "group":  "LSK",
        "level":  1,
        "totalCount":  9,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางเลสิค(Lasik Knowledge)",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LSK-19",
        "dept":  "GD",
        "title":  "Graphic Editor - LSK",
        "group":  "LSK",
        "level":  1,
        "totalCount":  9,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ด้านระบบคุณภาพ (Quality System Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การสื่อสาร (Communication)",
                             "การออกแบบและการจัดทำสื่อ (Graphic Design and Print media Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LSK-20",
        "dept":  "LSK",
        "title":  "Content Creator - LSK",
        "group":  "MKT",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางเลสิค(Lasik Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LSK-21",
        "dept":  "LSK",
        "title":  "Marketing Executive - LSK",
        "group":  "MKT",
        "level":  1,
        "totalCount":  10,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางเลสิค(Lasik Knowledge)",
                             "การแก้ไขปัญหาและการตัดสินใจ Problem Solving \u0026 Decision Making",
                             "การสื่อสาร (Communication)",
                             "การนำเสนอ(Presentation)",
                             "การบริหารการตลาด(Marketing Management)"
                         ]
    },
    {
        "jdCode":  "SD-HRD-01/LSK-22",
        "dept":  "LSK",
        "title":  "นักทัศนมาตร",
        "group":  "LSK",
        "level":  1,
        "totalCount":  9,
        "competencies":  [
                             "Relationship Excellence",
                             "Fact",
                             "Innovative Thinking",
                             "Learning",
                             "Entrepreneurship",
                             "ความรู้ทางเลสิค(Lasik Knowledge)",
                             "การปฏิบัติงานตามขั้นตอนและการจัดการเอกสาร(Daily \u0026 Documentation Management)",
                             "การสื่อสาร (Communication)",
                             "การบริการด้วยใจ(Service Mind)"
                         ]
    }
];

// --------------------------------------------------------------------------
// 3. CHART INSTANCES & SETUP (Light Mode Optimized Colors)
// --------------------------------------------------------------------------
let charts = {};

// ─── Tab titles for topbar ───
const TAB_TITLES = {
  overview: 'ภาพรวมและเป้าหมาย',
  jd:       'เกณฑ์ประเมินสมรรถนะ 124 ตำแหน่ง',
  core:     'Core Competency Analysis',
  job:      'Job Competency Analysis',
  matrix:   'Heatmap Matrix ทุกมิติ',
  idp:      'Individual Development Plan (IDP)'
};

document.addEventListener('DOMContentLoaded', () => {
  initTabNav();
  initCharts();
  renderCoreSummary();
  renderJDTable(JDPositionsData);
  renderIDPTable();
  initModalListeners();
  initJDSearch();
});

// ─── Tab navigation ───
function initTabNav() {
  const navBtns  = document.querySelectorAll('.nav-item[data-tab]');
  const panels   = document.querySelectorAll('.tab-panel');
  const titleEl  = document.getElementById('topbarTitle');

  navBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const tab = btn.dataset.tab;
      navBtns.forEach(b => b.classList.remove('active'));
      panels.forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      const panel = document.getElementById('tab-' + tab);
      if (panel) panel.classList.add('active');
      if (titleEl) titleEl.textContent = TAB_TITLES[tab] || '';
    });
  });
}

// ─── Charts ───
let charts = {};

function initCharts() {
  Chart.defaults.font.family = "'Kanit', sans-serif";
  Chart.defaults.color = '#64748b';

  // 1. Radar – Core Competency
  const cRadar = document.getElementById('chartRadar')?.getContext('2d');
  if (cRadar) {
    charts.radar = new Chart(cRadar, {
      type: 'radar',
      data: {
        labels: ['Relationship','Fact','Innovative','Learning','Entrepreneurship'],
        datasets: [
          {
            label: '% Gap ทั้งบริษัท',
            data: [6.41, 12.81, 25.62, 31.32, 20.64],
            backgroundColor: 'rgba(37,99,235,.15)',
            borderColor: '#2563eb',
            pointBackgroundColor: '#2563eb',
            borderWidth: 2.5,
            pointRadius: 4
          },
          {
            label: 'เป้าหมายสูงสุด 30%',
            data: [30,30,30,30,30],
            backgroundColor: 'transparent',
            borderColor: 'rgba(220,38,38,.4)',
            borderDash: [5,4],
            pointRadius: 0,
            borderWidth: 1.5
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          r: {
            angleLines: { color: '#e2e8f0' },
            grid: { color: '#e2e8f0' },
            pointLabels: { color: '#334155', font: { size: 11, weight: '600', family: "'Kanit'" } },
            ticks: { backdropColor: 'transparent', color: '#94a3b8', font: { size: 9 } }
          }
        },
        plugins: { legend: { position: 'bottom', labels: { font: { family: "'Kanit'" }, boxWidth: 12 } } }
      }
    });
  }

  // 2. Horizontal bar – Top 5 Job Gaps
  const cJob = document.getElementById('chartJobBar')?.getContext('2d');
  if (cJob) {
    charts.jobBar = new Chart(cJob, {
      type: 'bar',
      indexAxis: 'y',
      data: {
        labels: [
          'ความรู้ระบบคุณภาพ',
          'การวิเคราะห์ข้อมูล',
          'การสื่อสาร',
          'การแก้ไขปัญหา & ตัดสินใจ',
          'ปฏิบัติงานตามขั้นตอน & เอกสาร'
        ],
        datasets: [{
          label: '% Gap',
          data: [45.45, 31.25, 20.37, 17.23, 8.86],
          backgroundColor: ['#dc2626','#d97706','#2563eb','#0891b2','#16a34a'],
          borderRadius: 5,
          borderSkipped: false
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { grid: { color: '#f1f5f9' }, ticks: { callback: v => v + '%', font: { family: "'Kanit'" } } },
          y: { grid: { display: false }, ticks: { font: { family: "'Kanit'", size: 11 } } }
        },
        plugins: { legend: { display: false } }
      }
    });
  }

  // 3. Grouped bar – Core by Level
  const cLevel = document.getElementById('chartCoreLevel')?.getContext('2d');
  if (cLevel) {
    charts.coreLevel = new Chart(cLevel, {
      type: 'bar',
      data: {
        labels: ['Relationship','Fact','Innovative','Learning','Entrepreneurship'],
        datasets: [
          { label: 'พนักงาน (233)', data: [1.29,4.72,18.88,29.18,15.02], backgroundColor: '#0891b2', borderRadius: 3 },
          { label: 'หัวหน้างาน (34)', data: [23.53,44.12,58.82,38.24,50.00], backgroundColor: '#d97706', borderRadius: 3 },
          { label: 'ผู้จัดการ (14)', data: [50.00,71.43,57.14,50.00,42.86], backgroundColor: '#dc2626', borderRadius: 3 }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { grid: { display: false }, ticks: { font: { family: "'Kanit'", size: 10 } } },
          y: { grid: { color: '#f1f5f9' }, ticks: { callback: v => v + '%', font: { family: "'Kanit'" } } }
        },
        plugins: { legend: { position: 'bottom', labels: { font: { family: "'Kanit'" }, boxWidth: 12 } } }
      }
    });
  }
}

// ─── Core Summary Table ───
function renderCoreSummary() {
  const tbody = document.getElementById('tblCoreSummary');
  if (!tbody) return;
  const data = CompetencyData.coreCompetency.overall;
  tbody.innerHTML = data.map((item, i) => {
    const cls = item.gapPct >= 30 ? 'red' : item.gapPct >= 15 ? 'amber' : 'green';
    const barCls = item.gapPct >= 30 ? 'red' : item.gapPct >= 15 ? 'amber' : 'green';
    return `
      <tr>
        <td style="color:var(--text-400);font-weight:600;">${i+1}</td>
        <td><strong>${item.name}</strong></td>
        <td style="text-align:center;">${item.total}</td>
        <td style="text-align:center;"><span class="tag tag-${cls}">${item.gapCount} ท่าน</span></td>
        <td style="text-align:center;"><strong style="color:var(--${cls === 'red' ? 'red' : cls === 'amber' ? 'amber' : 'green'});">${item.gapPct}%</strong></td>
        <td style="min-width:180px;">
          <div class="bar-wrap">
            <div class="bar-bg"><div class="bar-fill ${barCls}" style="width:${item.gapPct}%"></div></div>
            <span class="bar-pct">${item.gapPct}%</span>
          </div>
        </td>
      </tr>`;
  }).join('');
}

// ─── JD Matrix Table ───
function renderJDTable(data) {
  const tbody = document.getElementById('tblJD');
  const countEl = document.getElementById('jdCount');
  if (countEl) countEl.textContent = data.length + ' ตำแหน่ง';
  if (!tbody) return;

  tbody.innerHTML = data.map(pos => {
    const lvlLabel = pos.level === 3 ? 'ผู้จัดการ' : pos.level === 2 ? 'หัวหน้างาน' : 'พนักงาน';
    const lvlTag   = pos.level === 3 ? 'tag-red' : pos.level === 2 ? 'tag-amber' : 'tag-green';
    const coreComp = pos.competencies.filter(c =>
      ['Relationship Excellence','Fact','Innovative Thinking','Learning','Entrepreneurship'].includes(c)
    );
    const knowledgeComp = pos.competencies.filter(c =>
      c.includes('Knowledge') || c.includes('ความรู้') || c.includes('Lasik') || c.includes('Legal')
    );
    const skillComp = pos.competencies.filter(c =>
      !coreComp.includes(c) && !knowledgeComp.includes(c)
    );
    const chips =
      coreComp.map(c => `<span class="chip chip-core">${c}</span>`).join('') +
      knowledgeComp.map(c => `<span class="chip chip-know">${c.replace(/ \(.*\)/, '')}</span>`).join('') +
      skillComp.map(c => `<span class="chip chip-skill">${c.replace(/ \(.*\)/, '').replace(/[\(\)].*$/, '').trim()}</span>`).join('');

    return `
      <tr>
        <td><code style="font-size:.7rem;color:var(--text-400);">${pos.jdCode}</code></td>
        <td><span class="tag tag-blue">${pos.dept}</span></td>
        <td><strong>${pos.title}</strong></td>
        <td><span class="tag ${lvlTag}">L${pos.level} · ${lvlLabel}</span></td>
        <td style="text-align:center;"><strong>${pos.totalCount}</strong></td>
        <td style="max-width:480px;line-height:1.8;">${chips}</td>
      </tr>`;
  }).join('');
}

// ─── JD Search & Filter ───
function initJDSearch() {
  const search = document.getElementById('searchJD');
  const dept   = document.getElementById('jdDept');
  const level  = document.getElementById('jdLevel');

  function doFilter() {
    const q   = search?.value.toLowerCase().trim() || '';
    const d   = dept?.value   || 'all';
    const lv  = level?.value  || 'all';
    const res = JDPositionsData.filter(p => {
      const mq = !q || p.title.toLowerCase().includes(q) ||
                       p.jdCode.toLowerCase().includes(q) ||
                       (p.dept||'').toLowerCase().includes(q);
      const md = d  === 'all' || p.dept === d;
      const ml = lv === 'all' || String(p.level) === lv;
      return mq && md && ml;
    });
    renderJDTable(res);
  }

  search?.addEventListener('input',  doFilter);
  dept?.addEventListener('change',   doFilter);
  level?.addEventListener('change',  doFilter);
}

// ─── IDP Table ───
function renderIDPTable() {
  const tbody = document.getElementById('tblIDP');
  if (!tbody) return;
  tbody.innerHTML = CompetencyData.idpActions.map(a => {
    const stCls = a.status === 'อนุมัติแล้ว' ? 'tag-green' :
                  a.status === 'กำลังดำเนินการ' ? 'tag-blue' : 'tag-amber';
    return `
      <tr>
        <td><code>${a.id}</code></td>
        <td><strong>${a.name}</strong></td>
        <td>${a.level}</td>
        <td>${a.group}</td>
        <td><span class="tag tag-red">${a.gapItem}</span></td>
        <td><span class="tag ${stCls}">${a.status}</span></td>
        <td>${a.dueDate}</td>
        <td><button class="btn btn-sm" onclick="alert('จัดการ IDP: ${a.name}')">จัดการ</button></td>
      </tr>`;
  }).join('');
}

// ─── Modal ───
function initModalListeners() {
  const modal   = document.getElementById('apiModal');
  const openBtn = document.getElementById('btnOpenApi');
  const closeBtn= document.getElementById('btnCloseApi');
  if (openBtn  && modal) openBtn.addEventListener('click',  () => modal.classList.add('open'));
  if (closeBtn && modal) closeBtn.addEventListener('click', () => modal.classList.remove('open'));
  modal?.addEventListener('click', e => { if (e.target === modal) modal.classList.remove('open'); });
}

// ─── Global API Service ───
window.CompetencyAPIService = {
  getSummary:    () => CompetencyData,
  getJDPositions:() => JDPositionsData,
  getCoreGaps:   () => CompetencyData.coreCompetency,
  getIDPActions: () => CompetencyData.idpActions
};

