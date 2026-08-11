import type { Student, Subject } from "@/lib/analysis";

export const meta = {
  institution: "Jeppiaar Nagar, Rajiv Gandhi Salai, Chennai - 119",
  department: "DEPARTMENT OF ARTIFICIAL INTELLIGENCE AND MACHINE LEARNING",
  title: "INTERNAL ASSESSMENT II RESULT ANALYSIS",
  year: "III",
  semester: "VI",
  batch: "2024-2028",
  section: "A",
  passMark: 50,
  maxMark: 100,
};

export const seedSubjects: Subject[] = [
  { code: "AL3501", name: "Natural Language Processing", staff: "Mr.B.N.Sairam AP/AIML" },
  { code: "AL3502", name: "Deep Learning for Vision", staff: "Mrs.P.S.Indhumathi AP/AIML" },
  { code: "CS3551", name: "Distributed Computing", staff: "Ms.S.Ilakkia AP/AIML" },
  { code: "CB3492", name: "Cryptography and Cyber Security", staff: "Dr.A.Vidhya HOD/AIML" },
  { code: "CCS335", name: "Cloud Computing", staff: "Mr.B.N.Sairam AP/AIML" },
  { code: "CCS366", name: "Software Testing Automation", staff: "Mrs.P.S.Indhumathi AP/AIML" },
];

export const seedStudents: Student[] = [
  { reg: "310824148001", name: "AARAV MENON", gender: "B", quota: "C", stay: "DS", marks: [92, 85, 88, 90, 78, 82] },
  { reg: "310824148002", name: "ABHINAYA R", gender: "G", quota: "C", stay: "H", marks: [78, 72, 80, 75, 81, 70] },
  { reg: "310824148003", name: "ADITYA KUMAR S", gender: "B", quota: "M", stay: "DS", marks: [45, 38, 52, 48, 55, 60] },
  { reg: "310824148004", name: "AISWARYA P", gender: "G", quota: "C", stay: "DS", marks: [95, 91, 93, 88, 90, 85] },
  { reg: "310824148005", name: "AKASH V", gender: "B", quota: "C", stay: "H", marks: [62, 55, 58, 70, 65, 72] },
  { reg: "310824148006", name: "ANANYA SRI", gender: "G", quota: "M", stay: "DS", marks: [88, 84, 79, 82, 86, 80] },
  { reg: "310824148007", name: "ARJUN RAJ", gender: "B", quota: "C", stay: "DS", marks: [30, 25, 40, "AB", 35, 42] },
  { reg: "310824148008", name: "BHAVYA N", gender: "G", quota: "C", stay: "H", marks: [75, 68, 72, 70, 74, 69] },
  { reg: "310824148009", name: "CHARAN M", gender: "B", quota: "M", stay: "H", marks: [55, 48, 60, 52, 58, 50] },
  { reg: "310824148010", name: "DEEPIKA R", gender: "G", quota: "C", stay: "DS", marks: [90, 87, 91, 85, 88, 84] },
  { reg: "310824148011", name: "DHANUSH K", gender: "B", quota: "M", stay: "DS", marks: [40, 35, 28, 42, 30, "AB"] },
  { reg: "310824148012", name: "DIVYA S", gender: "G", quota: "C", stay: "H", marks: [82, 78, 85, 80, 83, 76] },
  { reg: "310824148013", name: "GOKUL R", gender: "B", quota: "C", stay: "DS", marks: [68, 62, 70, 65, 72, 60] },
  { reg: "310824148014", name: "HARINI V", gender: "G", quota: "M", stay: "DS", marks: [93, 89, 95, 91, 90, 88] },
  { reg: "310824148015", name: "ISHAN P", gender: "B", quota: "C", stay: "H", marks: [50, 45, 55, 48, 52, 58] },
  { reg: "310824148016", name: "JANANI M", gender: "G", quota: "C", stay: "DS", marks: [77, 70, 75, 72, 78, 74] },
];
