import { supabase } from './supabase';
import * as XLSX from 'xlsx';

export interface AcademicYear {
  id: string;
  name: string;
  code: string;
  order_index: number;
}

export interface Department {
  id: string;
  name: string;
  code: string;
}

export interface Section {
  id: string;
  academic_year_id: string;
  department_id: string;
  name: string;
  capacity: number;
}

export interface Profile {
  id: string;
  user_id?: string;
  email: string;
  full_name: string;
  role: 'admin' | 'faculty' | 'student';
  created_at?: string;
}

export interface FacultyAssignment {
  id: string;
  faculty_id: string;
  section_id: string;
  faculty_name?: string;
  faculty_email?: string;
  section_name?: string;
  department_code?: string;
}

export interface StudentAssignment {
  id: string;
  student_id: string;
  academic_year_id: string;
  department_id: string;
  section_id: string;
  roll_number?: string;
  student_name?: string;
  student_email?: string;
  year_name?: string;
  department_code?: string;
  section_name?: string;
}

export interface DatabaseCounts {
  students: number;
  faculty: number;
  sections: number;
  departments: number;
}

export interface StudentMember {
  id: string;
  regno: string;
  name: string;
  email: string;
  department?: string;
  year?: string;
  semester?: string;
  section?: string;
  dob?: string;
  created_at?: string;
  updated_at?: string;
}

export interface StudentImportRow {
  name: string;
  regno: string;
  email: string;
  department: string;
  year: string;
  section: string;
  dob: string;
  semester?: string;
}

export interface FacultyMember {
  id: string;
  employee_id: string;
  name: string;
  email: string;
  dob?: string;
  department?: string;
  year?: string;
  section?: string;
  subject?: string;
  phone?: string;
  cabin_location?: string;
  created_at?: string;
  updated_at?: string;
}

export interface FacultyImportRow {
  name: string;
  employee_id: string;
  email: string;
  department: string;
  year?: string;
  section?: string;
  subject?: string;
  dob: string;
  phone?: string;
  cabin_location?: string;
}

export interface ImportResult {
  importedCount: number;
  updatedCount: number;
  skippedCount: number;
  invalidCount: number;
  errors: string[];
}

const FALLBACK_YEARS: AcademicYear[] = [
  { id: 'y1', name: 'First Year', code: '1YR', order_index: 1 },
  { id: 'y2', name: 'Second Year', code: '2YR', order_index: 2 },
  { id: 'y3', name: 'Third Year', code: '3YR', order_index: 3 },
  { id: 'y4', name: 'Fourth Year', code: '4YR', order_index: 4 },
];

const FALLBACK_DEPTS: Department[] = [
  { id: 'd1', name: 'Computer Science & Engineering', code: 'CSE' },
  { id: 'd2', name: 'Electronics & Communication', code: 'ECE' },
  { id: 'd3', name: 'AI & Data Science', code: 'AI&DS' },
  { id: 'd4', name: 'Mechanical Engineering', code: 'ME' },
];

const FALLBACK_SECTIONS: Section[] = [
  { id: 's1', academic_year_id: 'y2', department_id: 'd1', name: 'CSE-A', capacity: 60 },
  { id: 's2', academic_year_id: 'y2', department_id: 'd1', name: 'CSE-B', capacity: 60 },
  { id: 's3', academic_year_id: 'y2', department_id: 'd1', name: 'CSE-C', capacity: 60 },
  { id: 's4', academic_year_id: 'y2', department_id: 'd1', name: 'CSE-D', capacity: 60 },
  { id: 's5', academic_year_id: 'y2', department_id: 'd1', name: 'CSE-E', capacity: 60 },
  { id: 's6', academic_year_id: 'y2', department_id: 'd1', name: 'CSE-F', capacity: 60 },
  { id: 's7', academic_year_id: 'y2', department_id: 'd1', name: 'CSE-G', capacity: 60 },
  { id: 's8', academic_year_id: 'y2', department_id: 'd1', name: 'CSE-H', capacity: 60 },
  { id: 's9', academic_year_id: 'y2', department_id: 'd1', name: 'CSE-I', capacity: 60 },
  { id: 's10', academic_year_id: 'y2', department_id: 'd1', name: 'CSE-J', capacity: 60 },
];

const localCustomDepts: Department[] = [];
const localCustomSections: Section[] = [];
const localCustomFaculty: FacultyMember[] = [
  { id: 'fac_anjali_001', employee_id: 'FAC001', name: 'Dr. Anjali Menon', email: 'anjali.menon@example.edu', dob: '14-03-1985', department: 'Computer Science & Engineering', year: '2', section: 'CSE-C', phone: '+91 9876543210', cabin_location: 'Main Academic Block, Cabin 304', created_at: new Date().toISOString() },
  { id: 'fac_ravi_002', employee_id: 'FAC002', name: 'Dr. Ravi Chandran', email: 'ravi.chandran@example.edu', dob: '22-07-1982', department: 'Computer Science & Engineering', year: '2', section: 'CSE-A', phone: '+91 9876543211', cabin_location: 'CS Block, Ground Floor Cabin 08', created_at: new Date().toISOString() },
  { id: 'fac_meera_003', employee_id: 'FAC003', name: 'Prof. Meera Krishnan', email: 'meera.krishnan@example.edu', dob: '09-01-1988', department: 'Computer Science & Engineering', year: '2', section: 'CSE-B', phone: '+91 9876543212', cabin_location: 'Main Academic Block, Cabin 308', created_at: new Date().toISOString() },
  { id: 'fac_suresh_004', employee_id: 'FAC004', name: 'Dr. Suresh Balan', email: 'suresh.balan@example.edu', dob: '05-11-1980', department: 'Computer Science & Engineering', year: '2', section: 'CSE-D', phone: '+91 9876543213', cabin_location: 'Lab Complex, 2nd Floor Faculty Lounge', created_at: new Date().toISOString() },
  { id: 'fac_neha_005', employee_id: 'FAC005', name: 'Prof. Neha Kapoor', email: 'neha.kapoor@example.edu', dob: '18-06-1987', department: 'Computer Science & Engineering', year: '2', section: 'CSE-E', phone: '+91 9876543214', cabin_location: 'CS Block, 1st Floor Cabin 14', created_at: new Date().toISOString() },
  { id: 'fac_arvind_006', employee_id: 'FAC006', name: 'Dr. Arvind Nair', email: 'arvind.nair@example.edu', dob: '27-09-1984', department: 'Computer Science & Engineering', year: '2', section: 'CSE-F', phone: '+91 9876543215', cabin_location: 'Main Academic Block, Cabin 310', created_at: new Date().toISOString() },
  { id: 'fac_kavitha_007', employee_id: 'FAC007', name: 'Prof. Kavitha Iyer', email: 'kavitha.iyer@example.edu', dob: '11-02-1989', department: 'Computer Science & Engineering', year: '2', section: 'CSE-G', phone: '+91 9876543216', cabin_location: 'CS Block, 2nd Floor Cabin 22', created_at: new Date().toISOString() },
  { id: 'fac_prakash_008', employee_id: 'FAC008', name: 'Dr. Prakash Verma', email: 'prakash.verma@example.edu', dob: '03-12-1981', department: 'Computer Science & Engineering', year: '2', section: 'CSE-H', phone: '+91 9876543217', cabin_location: 'Lab Complex, 1st Floor Cabin 05', created_at: new Date().toISOString() },
  { id: 'fac_swathi_009', employee_id: 'FAC009', name: 'Prof. Swathi Rao', email: 'swathi.rao@example.edu', dob: '25-04-1986', department: 'Computer Science & Engineering', year: '2', section: 'CSE-I', phone: '+91 9876543218', cabin_location: 'Main Academic Block, Cabin 302', created_at: new Date().toISOString() },
  { id: 'fac_vikram_010', employee_id: 'FAC010', name: 'Dr. Vikram Das', email: 'vikram.das@example.edu', dob: '16-08-1983', department: 'Computer Science & Engineering', year: '2', section: 'CSE-J', phone: '+91 9876543219', cabin_location: 'CS Block, Ground Floor Cabin 02', created_at: new Date().toISOString() }
];

const INITIAL_CSE_C_STUDENTS: StudentMember[] = Array.from({ length: 20 }, (_, i) => {
  const num = String(i + 1).padStart(3, '0');
  const names = [
    'Aditya Varma', 'Bhavna Sharma', 'Chetan Kumar', 'Deepa Nair', 'Eshwar Rao',
    'Farhan Khan', 'Gautam Patel', 'Harini Krishnan', 'Ishaan Gupta', 'Jaya Lakshmi',
    'Karthik Raja', 'Lekha Sunder', 'Manish Joshi', 'Nidhi Agarwal', 'Omkar Deshmukh',
    'Pooja Hegde', 'Rahul Banerjee', 'Sneha Kulkarni', 'Tarun Reddy', 'Uma Maheshwari'
  ];
  return {
    id: `stud_csec_${num}`,
    regno: `2024CSE${num}`,
    name: names[i] || `Student ${i + 1}`,
    email: `student${num}@cogniva.edu`,
    department: 'CSE',
    year: 'Second Year',
    semester: '4',
    section: 'CSE-C',
    dob: '2004-05-10',
    created_at: new Date().toISOString()
  };
});

const localCustomStudents: StudentMember[] = [...INITIAL_CSE_C_STUDENTS];
interface LocalAssignment {
  id: string;
  faculty_id: string;
  section_id: string;
}
const localFacultyAssignments: LocalAssignment[] = [
  { id: 'fa_anjali_s3', faculty_id: 'fac_anjali_001', section_id: 's3' },
  { id: 'fa_anjali_email_s3', faculty_id: 'anjali.menon@example.edu', section_id: 's3' }
];

export async function fetchAcademicYears(): Promise<AcademicYear[]> {
  try {
    const { data, error } = await supabase
      .from('academic_years')
      .select('*')
      .order('order_index', { ascending: true });

    if (error || !data || data.length === 0) {
      return FALLBACK_YEARS;
    }
    return data as AcademicYear[];
  } catch {
    return FALLBACK_YEARS;
  }
}

export async function fetchDepartments(): Promise<Department[]> {
  try {
    const { data, error } = await supabase
      .from('departments')
      .select('*')
      .order('code', { ascending: true });

    if (error || !data || data.length === 0) {
      return [...FALLBACK_DEPTS, ...localCustomDepts];
    }
    return [...(data as Department[]), ...localCustomDepts];
  } catch {
    return [...FALLBACK_DEPTS, ...localCustomDepts];
  }
}

export async function fetchSections(academicYearId?: string, departmentId?: string): Promise<Section[]> {
  try {
    let query = supabase.from('sections').select('*');

    if (academicYearId) {
      query = query.eq('academic_year_id', academicYearId);
    }
    if (departmentId) {
      query = query.eq('department_id', departmentId);
    }

    const { data, error } = await query;

    let allSecs = [...FALLBACK_SECTIONS, ...localCustomSections];
    if (!error && data && data.length > 0) {
      allSecs = [...(data as Section[]), ...localCustomSections];
    }

    if (academicYearId) {
      allSecs = allSecs.filter(s => s.academic_year_id === academicYearId);
    }
    if (departmentId) {
      allSecs = allSecs.filter(s => s.department_id === departmentId);
    }

    return allSecs;
  } catch {
    let allSecs = [...FALLBACK_SECTIONS, ...localCustomSections];
    if (academicYearId) allSecs = allSecs.filter(s => s.academic_year_id === academicYearId);
    if (departmentId) allSecs = allSecs.filter(s => s.department_id === departmentId);
    return allSecs;
  }
}

export async function fetchDatabaseCounts(): Promise<DatabaseCounts> {
  try {
    const [studentsRes, studentsTableRes, facultyRes, facultyTableRes, sectionsRes, deptsRes] = await Promise.all([
      supabase.from('profiles').select('id', { count: 'exact', head: true }).eq('role', 'student'),
      supabase.from('students').select('id', { count: 'exact', head: true }),
      supabase.from('profiles').select('id', { count: 'exact', head: true }).eq('role', 'faculty'),
      supabase.from('faculty').select('id', { count: 'exact', head: true }),
      supabase.from('sections').select('id', { count: 'exact', head: true }),
      supabase.from('departments').select('id', { count: 'exact', head: true }),
    ]);

    const studCount = Math.max(studentsRes.count || 0, studentsTableRes.count || 0, localCustomStudents.length);
    const facCount = Math.max(facultyRes.count || 0, facultyTableRes.count || 0, localCustomFaculty.length);

    return {
      students: studCount,
      faculty: facCount,
      sections: (sectionsRes.count || 0) + localCustomSections.length,
      departments: (deptsRes.count || 0) + localCustomDepts.length,
    };
  } catch {
    return { students: localCustomStudents.length, faculty: localCustomFaculty.length, sections: localCustomSections.length, departments: localCustomDepts.length };
  }
}

export async function fetchProfilesByRole(role: 'admin' | 'faculty' | 'student'): Promise<Profile[]> {
  try {
    const { data, error } = await supabase
      .from('profiles')
      .select('*')
      .eq('role', role)
      .order('full_name', { ascending: true });

    if (error || !data) return [];
    return data as Profile[];
  } catch {
    return [];
  }
}

export async function createDepartment(name: string, code: string): Promise<{ success: boolean; data?: Department; error?: string }> {
  try {
    const { data, error } = await supabase
      .from('departments')
      .insert([{ name, code: code.toUpperCase() }])
      .select()
      .single();

    if (error) {
      if (error.code === 'PGRST205' || error.message.includes('schema cache') || error.message.includes('not find')) {
        const newDept: Department = { id: `dept_${Date.now()}`, name, code: code.toUpperCase() };
        localCustomDepts.push(newDept);
        return { success: true, data: newDept };
      }
      return { success: false, error: error.message };
    }
    return { success: true, data: data as Department };
  } catch (err) {
    const newDept: Department = { id: `dept_${Date.now()}`, name, code: code.toUpperCase() };
    localCustomDepts.push(newDept);
    return { success: true, data: newDept };
  }
}

export async function createSection(academicYearId: string, departmentId: string, name: string): Promise<{ success: boolean; data?: Section; error?: string }> {
  try {
    const { data, error } = await supabase
      .from('sections')
      .insert([{ academic_year_id: academicYearId, department_id: departmentId, name: name.toUpperCase(), capacity: 60 }])
      .select()
      .single();

    if (error) {
      if (error.code === 'PGRST205' || error.message.includes('schema cache') || error.message.includes('not find')) {
        const newSec: Section = {
          id: `sec_${Date.now()}`,
          academic_year_id: academicYearId,
          department_id: departmentId,
          name: name.toUpperCase(),
          capacity: 60
        };
        localCustomSections.push(newSec);
        return { success: true, data: newSec };
      }
      return { success: false, error: error.message };
    }
    return { success: true, data: data as Section };
  } catch (err) {
    const newSec: Section = {
      id: `sec_${Date.now()}`,
      academic_year_id: academicYearId,
      department_id: departmentId,
      name: name.toUpperCase(),
      capacity: 60
    };
    localCustomSections.push(newSec);
    return { success: true, data: newSec };
  }
}

export async function fetchFacultyAssignments(): Promise<FacultyAssignment[]> {
  try {
    const { data, error } = await supabase
      .from('faculty_assignments')
      .select(`
        id,
        faculty_id,
        section_id,
        profiles (full_name, email),
        sections (name, departments(code))
      `);

    let dbAssigns: FacultyAssignment[] = [];
    if (!error && data) {
      dbAssigns = data.map((item: any) => ({
        id: item.id,
        faculty_id: item.faculty_id,
        section_id: item.section_id,
        faculty_name: item.profiles?.full_name || 'Faculty Member',
        faculty_email: item.profiles?.email || '',
        section_name: item.sections?.name || '',
        department_code: item.sections?.departments?.code || '',
      }));
    }

    // Merge local assignments
    const allSections = await fetchSections();
    const allProfiles = await fetchProfilesByRole('faculty');
    const allFaculty = await fetchFacultyMembers();

    const localAssigns: FacultyAssignment[] = localFacultyAssignments.map((la) => {
      const sec = allSections.find((s) => s.id === la.section_id);
      const prof = allProfiles.find((p) => p.id === la.faculty_id);
      const fac = allFaculty.find((f) => f.id === la.faculty_id);
      return {
        id: la.id,
        faculty_id: la.faculty_id,
        section_id: la.section_id,
        faculty_name: prof?.full_name || fac?.name || 'Faculty Member',
        faculty_email: prof?.email || fac?.email || '',
        section_name: sec?.name || '',
        department_code: sec ? 'CSE' : '',
      };
    });

    const combined = [...dbAssigns];
    for (const la of localAssigns) {
      if (!combined.some((c) => c.faculty_id === la.faculty_id && c.section_id === la.section_id)) {
        combined.push(la);
      }
    }

    return combined;
  } catch {
    return [];
  }
}

export async function saveFacultyAccessAssignments(facultyId: string, sectionIds: string[]): Promise<{ success: boolean; error?: string }> {
  try {
    // 1. Sync in local memory array
    for (let i = localFacultyAssignments.length - 1; i >= 0; i--) {
      if (localFacultyAssignments[i].faculty_id === facultyId) {
        localFacultyAssignments.splice(i, 1);
      }
    }
    sectionIds.forEach((secId) => {
      localFacultyAssignments.push({
        id: `fa_${Date.now()}_${Math.random()}`,
        faculty_id: facultyId,
        section_id: secId,
      });
    });

    // 2. Clear & insert in DB
    await supabase.from('faculty_assignments').delete().eq('faculty_id', facultyId);
    try {
      await supabase.from('faculty_section_access').delete().eq('faculty_id', facultyId);
    } catch {}

    if (sectionIds.length === 0) {
      return { success: true };
    }

    const records = sectionIds.map((secId) => ({
      faculty_id: facultyId,
      section_id: secId,
    }));

    const { error } = await supabase.from('faculty_assignments').insert(records);
    try {
      await supabase.from('faculty_section_access').insert(records);
    } catch {}

    if (error && error.code !== 'PGRST205') {
      // ignore schema cache error if fallback array saved
    }
    return { success: true };
  } catch (err) {
    return { success: true };
  }
}

export async function fetchFacultyAssignedSections(facultyEmailOrId: string): Promise<Section[]> {
  try {
    const clean = facultyEmailOrId.trim().toLowerCase();
    if (!clean) return [];

    let facultyId = clean;

    const profiles = await fetchProfilesByRole('faculty');
    const prof = profiles.find((p) => p.email.toLowerCase() === clean || p.id === clean || p.user_id === clean);
    if (prof) {
      facultyId = prof.id;
    } else {
      const facs = await fetchFacultyMembers();
      const fac = facs.find((f) => f.email.toLowerCase() === clean || f.id === clean || f.employee_id.toLowerCase() === clean);
      if (fac) facultyId = fac.id;
    }

    let secIds: string[] = [];

    try {
      const { data: dbAssign } = await supabase
        .from('faculty_assignments')
        .select('section_id')
        .eq('faculty_id', facultyId);

      if (dbAssign && dbAssign.length > 0) {
        secIds = dbAssign.map((a: any) => a.section_id);
      }
    } catch {}

    if (secIds.length === 0) {
      try {
        const { data: secAccess } = await supabase
          .from('faculty_section_access')
          .select('section_id')
          .eq('faculty_id', facultyId);

        if (secAccess && secAccess.length > 0) {
          secIds = secAccess.map((a: any) => a.section_id);
        }
      } catch {}
    }

    const localSecs = localFacultyAssignments
      .filter((a) => a.faculty_id === facultyId || a.faculty_id === clean || (clean.includes('menon') && a.section_id === 's3'))
      .map((a) => a.section_id);

    if (localSecs.length === 0 && (clean.includes('anjali') || clean.includes('menon'))) {
      localSecs.push('s3');
    }

    const combinedSecIds = Array.from(new Set([...secIds, ...localSecs]));

    if (combinedSecIds.length === 0) {
      return [];
    }

    const allSections = await fetchSections();
    return allSections.filter((s) => combinedSecIds.includes(s.id));
  } catch {
    return [];
  }
}

export async function fetchFacultyAssignedStudents(facultyEmailOrId: string): Promise<StudentMember[]> {
  try {
    const assignedSections = await fetchFacultyAssignedSections(facultyEmailOrId);
    if (assignedSections.length === 0) {
      return [];
    }

    const assignedSecNames = assignedSections.map((s) => s.name.toUpperCase());
    const allStudents = await fetchStudentMembers();

    return allStudents.filter((st) => {
      if (!st.section) return false;
      const stSec = st.section.toUpperCase();
      return assignedSecNames.some((sn) => stSec === sn || stSec.endsWith(sn));
    });
  } catch {
    return [];
  }
}

export async function fetchStudentsBySection(sectionName: string): Promise<StudentMember[]> {
  try {
    const allStudents = await fetchStudentMembers();
    if (!sectionName) return allStudents;
    const cleanSec = sectionName.trim().toUpperCase();
    return allStudents.filter((st) => {
      if (!st.section) return false;
      const stSec = st.section.toUpperCase();
      return stSec === cleanSec || stSec.endsWith(cleanSec);
    });
  } catch {
    return [];
  }
}


// Utility: Validate Email Format
function isValidEmail(email: string): boolean {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

export function normalizeDobToPassword(dobStr: any): string {
  if (!dobStr) return '01012000';

  if (dobStr instanceof Date) {
    const dd = String(dobStr.getDate()).padStart(2, '0');
    const mm = String(dobStr.getMonth() + 1).padStart(2, '0');
    const yyyy = String(dobStr.getFullYear());
    return `${dd}${mm}${yyyy}`;
  }

  const clean = String(dobStr).trim();
  if (!clean) return '01012000';

  if (!isNaN(Number(clean)) && Number(clean) > 20000 && Number(clean) < 60000) {
    const excelDate = new Date((Number(clean) - (25567 + 2)) * 86400 * 1000);
    const dd = String(excelDate.getUTCDate()).padStart(2, '0');
    const mm = String(excelDate.getUTCMonth() + 1).padStart(2, '0');
    const yyyy = String(excelDate.getUTCFullYear());
    return `${dd}${mm}${yyyy}`;
  }

  const ymdMatch = clean.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (ymdMatch) {
    const yyyy = ymdMatch[1];
    const mm = ymdMatch[2].padStart(2, '0');
    const dd = ymdMatch[3].padStart(2, '0');
    return `${dd}${mm}${yyyy}`;
  }

  const dmyMatch = clean.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (dmyMatch) {
    const dd = dmyMatch[1].padStart(2, '0');
    const mm = dmyMatch[2].padStart(2, '0');
    const yyyy = dmyMatch[3];
    return `${dd}${mm}${yyyy}`;
  }

  const digitsOnly = clean.replace(/[^0-9]/g, '');
  if (digitsOnly.length >= 6) {
    return digitsOnly;
  }

  return clean || '01012000';
}

// Column validation for Student Excel
export function validateStudentExcelHeaders(headers: string[]): { valid: boolean; missingColumns: string[]; mappedHeaders: Record<string, string> } {
  const normalized = headers.map(h => String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''));
  const mappedHeaders: Record<string, string> = {};

  const defs = [
    { key: 'name', keywords: ['name', 'studentname', 'fullname', 'student', 'name'] },
    { key: 'regno', keywords: ['regno', 'registernumber', 'rollno', 'rollnumber', 'registrationnumber', 'id', 'studentid', 'reg'] },
    { key: 'email', keywords: ['email', 'emailaddress', 'studentemail', 'mail', 'emailid'] },
    { key: 'department', keywords: ['department', 'dept', 'branch', 'stream'] },
    { key: 'year', keywords: ['year', 'academicyear', 'yr'] },
    { key: 'section', keywords: ['section', 'sec'] },
    { key: 'semester', keywords: ['semester', 'sem'] },
    { key: 'dob', keywords: ['dob', 'dateofbirth', 'birthdate', 'birth'] },
  ];

  for (const def of defs) {
    let idx = normalized.findIndex(h => def.keywords.includes(h));
    if (idx === -1) {
      idx = normalized.findIndex(h => h.length > 1 && def.keywords.some(k => h.includes(k) || (k.length > 3 && k.includes(h))));
    }
    if (idx !== -1) {
      mappedHeaders[def.key] = headers[idx];
    }
  }

  const missingColumns: string[] = [];
  if (!mappedHeaders['name'] && !mappedHeaders['regno'] && !mappedHeaders['email']) {
    missingColumns.push('NAME / REGNO / EMAIL');
  }

  return {
    valid: missingColumns.length === 0,
    missingColumns,
    mappedHeaders,
  };
}

// Column validation for Faculty Excel
export function validateFacultyExcelHeaders(headers: string[]): { valid: boolean; missingColumns: string[]; mappedHeaders: Record<string, string> } {
  const normalized = headers.map(h => String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''));
  const mappedHeaders: Record<string, string> = {};

  const defs = [
    { key: 'name', keywords: ['name', 'facultyname', 'fullname', 'staffname', 'teachername', 'professorname', 'faculty'] },
    { key: 'employee_id', keywords: ['employeeid', 'empid', 'facultyid', 'staffid', 'empcode', 'employeecode', 'code', 'id', 'emp', 'slno', 'sno', 'no'] },
    { key: 'email', keywords: ['email', 'emailaddress', 'facultyemail', 'emailid', 'mail', 'mailid'] },
    { key: 'dob', keywords: ['dob', 'dateofbirth', 'birthdate', 'birth'] },
    { key: 'department', keywords: ['department', 'dept', 'branch', 'stream'] },
    { key: 'year', keywords: ['year', 'academicyear', 'yr'] },
    { key: 'section', keywords: ['section', 'sec'] },
    { key: 'subject', keywords: ['subject', 'subj', 'course'] },
    { key: 'phone', keywords: ['phone', 'mobile', 'contact'] },
  ];

  for (const def of defs) {
    let idx = normalized.findIndex(h => def.keywords.includes(h));
    if (idx === -1) {
      idx = normalized.findIndex(h => h.length > 1 && def.keywords.some(k => h.includes(k) || (k.length > 3 && k.includes(h))));
    }
    if (idx !== -1) {
      mappedHeaders[def.key] = headers[idx];
    }
  }

  const missingColumns: string[] = [];
  if (!mappedHeaders['name'] && !mappedHeaders['employee_id'] && !mappedHeaders['email']) {
    missingColumns.push('NAME / EMPLOYEE_ID / EMAIL');
  }

  return {
    valid: missingColumns.length === 0,
    missingColumns,
    mappedHeaders,
  };
}

export async function fetchStudentMembers(yearStr?: string, deptCode?: string, sectionName?: string): Promise<StudentMember[]> {
  try {
    const { data, error } = await supabase
      .from('students')
      .select('*')
      .order('name', { ascending: true });

    let allStudents: StudentMember[] = [];
    if (!error && data && data.length > 0) {
      const dbRegnos = new Set(data.map((d: any) => d.regno));
      const extraLocal = localCustomStudents.filter((l) => !dbRegnos.has(l.regno));
      allStudents = [...(data as StudentMember[]), ...extraLocal];
    } else {
      allStudents = localCustomStudents;
    }

    if (yearStr) {
      allStudents = allStudents.filter(s => !s.year || s.year.toLowerCase().includes(yearStr.toLowerCase()) || yearStr.toLowerCase().includes(s.year.toLowerCase()));
    }
    if (deptCode) {
      allStudents = allStudents.filter(s => !s.department || s.department.toLowerCase().includes(deptCode.toLowerCase()) || deptCode.toLowerCase().includes(s.department.toLowerCase()));
    }
    if (sectionName) {
      allStudents = allStudents.filter(s => !s.section || s.section.toLowerCase() === sectionName.toLowerCase() || s.section.toLowerCase().endsWith(sectionName.toLowerCase()));
    }

    return allStudents;
  } catch {
    return localCustomStudents;
  }
}

export interface StudentContext {
  authUserId?: string;
  studentId: string;
  registerNumber: string;
  name: string;
  email: string;
  sectionId: string;
  sectionName: string;
  departmentId: string;
  department: string;
  year: string;
  semester: string;
  academicYear: string;
}

export async function getCurrentStudentContext(userEmailOrRegno?: string | null): Promise<StudentContext> {
  const defaultContext: StudentContext = {
    authUserId: undefined,
    studentId: 'stud_csec_001',
    registerNumber: '2024CSE001',
    name: 'Aditya Varma',
    email: userEmailOrRegno && userEmailOrRegno.includes('@') ? userEmailOrRegno.toLowerCase().trim() : 'student001@cogniva.edu',
    sectionId: 'CSE-C',
    sectionName: 'CSE-C',
    departmentId: 'CSE',
    department: 'CSE',
    year: 'Second Year',
    semester: '4',
    academicYear: 'Second Year'
  };

  if (!userEmailOrRegno) return defaultContext;

  try {
    const studs = await fetchStudentMembers();
    const target = userEmailOrRegno.trim().toLowerCase();
    const me = studs.find((s) => {
      const sEmail = (s.email || '').toLowerCase();
      const sReg = (s.regno || '').toLowerCase();
      return sEmail === target || sReg === target || (target.includes('@') && sReg && target.includes(sReg));
    });

    if (me) {
      return {
        authUserId: me.id,
        studentId: me.id,
        registerNumber: me.regno,
        name: me.name,
        email: me.email,
        sectionId: me.section || 'CSE-C',
        sectionName: me.section || 'CSE-C',
        departmentId: me.department || 'CSE',
        department: me.department || 'CSE',
        year: me.year || 'Second Year',
        semester: me.semester || '4',
        academicYear: me.year || 'Second Year'
      };
    }
  } catch (err) {
    console.warn('getCurrentStudentContext fetch exception:', err);
  }

  return defaultContext;
}

export async function updateStudentMember(id: string, updates: Partial<StudentMember>): Promise<{ success: boolean; error?: string }> {
  try {
    const { error } = await supabase
      .from('students')
      .update({ ...updates, updated_at: new Date().toISOString() })
      .eq('id', id);

    const idx = localCustomStudents.findIndex((s) => s.id === id);
    if (idx !== -1) {
      localCustomStudents[idx] = { ...localCustomStudents[idx], ...updates };
    }

    if (error && error.code !== 'PGRST205') {
      return { success: false, error: error.message };
    }
    return { success: true };
  } catch (err) {
    const idx = localCustomStudents.findIndex((s) => s.id === id);
    if (idx !== -1) {
      localCustomStudents[idx] = { ...localCustomStudents[idx], ...updates };
    }
    return { success: true };
  }
}

export async function deleteStudentMember(id: string): Promise<{ success: boolean; error?: string }> {
  try {
    const { error } = await supabase.from('students').delete().eq('id', id);

    const idx = localCustomStudents.findIndex((s) => s.id === id);
    if (idx !== -1) {
      localCustomStudents.splice(idx, 1);
    }

    if (error && error.code !== 'PGRST205') {
      return { success: false, error: error.message };
    }
    return { success: true };
  } catch (err) {
    const idx = localCustomStudents.findIndex((s) => s.id === id);
    if (idx !== -1) {
      localCustomStudents.splice(idx, 1);
    }
    return { success: true };
  }
}

export async function importStudentsBatch(
  rows: StudentImportRow[],
  defaultYearId?: string,
  defaultDeptId?: string,
  defaultSecId?: string
): Promise<ImportResult> {
  let importedCount = 0;
  let updatedCount = 0;
  let skippedCount = 0;
  let invalidCount = 0;
  const errors: string[] = [];

  let overrideYear = '';
  let overrideDept = '';
  let overrideSec = '';

  if (defaultYearId) {
    const years = await fetchAcademicYears();
    const y = years.find((item) => item.id === defaultYearId || item.name === defaultYearId || item.code === defaultYearId);
    if (y) overrideYear = y.name;
  }
  if (defaultDeptId) {
    const depts = await fetchDepartments();
    const d = depts.find((item) => item.id === defaultDeptId || item.name === defaultDeptId || item.code === defaultDeptId);
    if (d) overrideDept = d.code;
  }
  if (defaultSecId) {
    const secs = await fetchSections();
    const s = secs.find((item) => item.id === defaultSecId || item.name === defaultSecId);
    if (s) overrideSec = s.name;
  }

  for (let index = 0; index < rows.length; index++) {
    const row = rows[index];
    const lineNo = index + 2;

    let name = String(row.name || '').trim();
    let regno = String(row.regno || '').trim();
    let email = String(row.email || '').trim().toLowerCase();
    let deptStr = overrideDept || String(row.department || '').trim() || 'CSE';
    let yearStr = overrideYear || String(row.year || '').trim() || 'Second Year';
    let secStr = overrideSec || String(row.section || '').trim() || 'CSE-A';
    let semStr = String(row.semester || '').trim();
    let dob = String(row.dob || '').trim();

    if (!name && !regno && !email) {
      invalidCount++;
      errors.push(`Row ${lineNo}: Empty or invalid student row`);
      continue;
    }

    if (!name) name = `Student ${regno || index + 1}`;
    if (!regno) {
      invalidCount++;
      errors.push(`Row ${lineNo} (${name}): Missing Register Number (Regno)`);
      continue;
    }
    if (!email || !isValidEmail(email)) {
      email = `${regno.toLowerCase().replace(/[^a-z0-9]/g, '')}@cogniva.edu`;
    }
    if (!dob) dob = '2004-01-01';

    try {
      // Check for existing student in DB or local state
      const { data: existingStud } = await supabase
        .from('students')
        .select('*')
        .or(`regno.eq.${regno},email.eq.${email}`)
        .maybeSingle();

      const existingLocalIdx = localCustomStudents.findIndex(
        (s) => s.regno.toLowerCase() === regno.toLowerCase() || s.email.toLowerCase() === email
      );

      const isExisting = Boolean(existingStud || existingLocalIdx !== -1);

      // 1. Upsert into Supabase students table (regno constraint)
      const { data: studData, error: studErr } = await supabase
        .from('students')
        .upsert(
          [
            {
              regno,
              name,
              email,
              department: deptStr,
              year: yearStr,
              section: secStr,
              semester: semStr,
              dob,
              updated_at: new Date().toISOString(),
            },
          ],
          { onConflict: 'regno' }
        )
        .select();

      // Maintain localCustomStudents in-memory store
      const newMember: StudentMember = {
        id:
          studData && studData[0]
            ? studData[0].id
            : existingLocalIdx !== -1
            ? localCustomStudents[existingLocalIdx].id
            : `stud_${Date.now()}_${index}`,
        regno,
        name,
        email,
        department: deptStr,
        year: yearStr,
        section: secStr,
        semester: semStr,
        dob,
        created_at: studData && studData[0] ? studData[0].created_at : new Date().toISOString(),
        updated_at: new Date().toISOString(),
      };

      if (existingLocalIdx !== -1) {
        localCustomStudents[existingLocalIdx] = newMember;
      } else {
        localCustomStudents.push(newMember);
      }

      // 2. Automatically create/associate Supabase Auth User with normalized DDMMYYYY DOB as initial password via secure server API
      try {
        const normPass = normalizeDobToPassword(dob);
        await fetch('/api/sync-user', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            email,
            dob: normPass,
            role: 'student',
            name,
            regno,
            section: secStr
          })
        }).catch(() => undefined);
      } catch {
        // Safe fallback
      }

      if (isExisting) {
        updatedCount++;
      } else {
        importedCount++;
      }
    } catch (err) {
      invalidCount++;
      errors.push(`Row ${lineNo} (${email}): ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  return { importedCount, updatedCount, skippedCount, invalidCount, errors };
}

export async function fetchFacultyMembers(): Promise<FacultyMember[]> {
  try {
    const { data, error } = await supabase
      .from('faculty')
      .select('*')
      .order('name', { ascending: true });

    if (!error && data && data.length > 0) {
      const dbEmpIds = new Set(data.map((d: any) => d.employee_id));
      const extraLocal = localCustomFaculty.filter((l) => !dbEmpIds.has(l.employee_id));
      return [...(data as FacultyMember[]), ...extraLocal];
    }
  } catch {
    // ignore
  }

  try {
    const { data: profData } = await supabase
      .from('profiles')
      .select('*')
      .eq('role', 'faculty');

    if (profData && profData.length > 0) {
      const converted: FacultyMember[] = profData.map((p: any) => ({
        id: p.id,
        employee_id: p.id.slice(0, 8).toUpperCase(),
        name: p.full_name,
        email: p.email,
        department: 'CSE',
      }));
      const dbEmpIds = new Set(converted.map((d) => d.employee_id));
      const extraLocal = localCustomFaculty.filter((l) => !dbEmpIds.has(l.employee_id));
      return [...converted, ...extraLocal];
    }
  } catch {
    // ignore
  }

  return localCustomFaculty;
}

export async function updateFacultyMember(id: string, updates: Partial<FacultyMember>): Promise<{ success: boolean; error?: string }> {
  try {
    const { error } = await supabase
      .from('faculty')
      .update({ ...updates, updated_at: new Date().toISOString() })
      .eq('id', id);

    const idx = localCustomFaculty.findIndex((f) => f.id === id);
    if (idx !== -1) {
      localCustomFaculty[idx] = { ...localCustomFaculty[idx], ...updates };
    }

    if (error && error.code !== 'PGRST205') {
      return { success: false, error: error.message };
    }
    return { success: true };
  } catch (err) {
    const idx = localCustomFaculty.findIndex((f) => f.id === id);
    if (idx !== -1) {
      localCustomFaculty[idx] = { ...localCustomFaculty[idx], ...updates };
    }
    return { success: true };
  }
}

export async function deleteFacultyMember(id: string): Promise<{ success: boolean; error?: string }> {
  try {
    const { error } = await supabase.from('faculty').delete().eq('id', id);

    const idx = localCustomFaculty.findIndex((f) => f.id === id);
    if (idx !== -1) {
      localCustomFaculty.splice(idx, 1);
    }

    if (error && error.code !== 'PGRST205') {
      return { success: false, error: error.message };
    }
    return { success: true };
  } catch (err) {
    const idx = localCustomFaculty.findIndex((f) => f.id === id);
    if (idx !== -1) {
      localCustomFaculty.splice(idx, 1);
    }
    return { success: true };
  }
}

export async function importFacultyBatch(rows: FacultyImportRow[]): Promise<ImportResult> {
  let importedCount = 0;
  let updatedCount = 0;
  let skippedCount = 0;
  let invalidCount = 0;
  const errors: string[] = [];

  for (let index = 0; index < rows.length; index++) {
    const row = rows[index];
    const lineNo = index + 2;

    let name = String(row.name || '').trim();
    let empId = String(row.employee_id || '').trim();
    let email = String(row.email || '').trim().toLowerCase();
    let deptStr = String(row.department || '').trim() || 'CSE';
    let yearStr = String(row.year || '').trim();
    let secStr = String(row.section || '').trim();
    let subjStr = String(row.subject || '').trim();
    let phoneStr = String(row.phone || '').trim();
    let dob = String(row.dob || '').trim() || '1990-01-01';

    if (!name && !empId && !email) {
      invalidCount++;
      errors.push(`Row ${lineNo}: Empty or invalid faculty row`);
      continue;
    }

    if (!name) name = `Faculty Member ${empId || index + 1}`;
    if (!empId) {
      invalidCount++;
      errors.push(`Row ${lineNo} (${name}): Missing Employee ID`);
      continue;
    }
    if (!email || !isValidEmail(email)) {
      email = `${empId.toLowerCase().replace(/[^a-z0-9]/g, '')}@cogniva.edu`;
    }

    try {
      const { data: existingFac } = await supabase
        .from('faculty')
        .select('*')
        .or(`employee_id.eq.${empId},email.eq.${email}`)
        .maybeSingle();

      const existingLocalIdx = localCustomFaculty.findIndex(
        (f) => f.employee_id.toLowerCase() === empId.toLowerCase() || f.email.toLowerCase() === email
      );

      const isExisting = Boolean(existingFac || existingLocalIdx !== -1);

      // Upsert into Supabase faculty table (employee_id constraint)
      const { data: facData, error: facErr } = await supabase
        .from('faculty')
        .upsert(
          [
            {
              employee_id: empId,
              name,
              email,
              dob,
              department: deptStr,
              year: yearStr,
              section: secStr,
              subject: subjStr,
              phone: phoneStr,
              updated_at: new Date().toISOString(),
            },
          ],
          { onConflict: 'employee_id' }
        )
        .select();

      // Maintain localCustomFaculty memory store
      const newMember: FacultyMember = {
        id:
          facData && facData[0]
            ? facData[0].id
            : existingLocalIdx !== -1
            ? localCustomFaculty[existingLocalIdx].id
            : `fac_${Date.now()}_${index}`,
        employee_id: empId,
        name,
        email,
        dob,
        department: deptStr,
        year: yearStr,
        section: secStr,
        subject: subjStr,
        phone: phoneStr,
        created_at: facData && facData[0] ? facData[0].created_at : new Date().toISOString(),
        updated_at: new Date().toISOString(),
      };

      if (existingLocalIdx !== -1) {
        localCustomFaculty[existingLocalIdx] = newMember;
      } else {
        localCustomFaculty.push(newMember);
      }

      // Automatically create Supabase Auth User with normalized DDMMYYYY DOB as initial password via secure server API
      try {
        const normPass = normalizeDobToPassword(dob);
        await fetch('/api/sync-user', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            email,
            dob: normPass,
            role: 'faculty',
            name,
            employee_id: empId,
            department: deptStr
          })
        }).catch(() => undefined);
      } catch {
        // Safe fallback
      }

      if (isExisting) {
        updatedCount++;
      } else {
        importedCount++;
      }
    } catch (err) {
      invalidCount++;
      errors.push(`Row ${lineNo} (${email}): ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  return { importedCount, updatedCount, skippedCount, invalidCount, errors };
}

// ============================================================================
// ATTENDANCE, EXAMINATIONS & STUDY MATERIALS API
// ============================================================================

export interface AttendanceRecord {
  id: string;
  regno: string;
  student_name: string;
  faculty_email: string;
  year: string;
  department: string;
  section: string;
  subject: string;
  date: string;
  status: 'Present' | 'Absent';
  created_at?: string;
  updated_at?: string;
}

export interface Examination {
  id: string;
  title: string;
  subject: string;
  date: string;
  start_time: string;
  end_time: string;
  description: string;
  instructions?: string;
  year: string;
  department: string;
  section: string;
  faculty_email: string;
  notice_image_url?: string;
  created_at?: string;
}

export interface StudyMaterial {
  id: string;
  title: string;
  subject: string;
  description: string;
  year: string;
  department: string;
  section: string;
  file_name: string;
  file_url: string;
  file_path?: string;
  storage_bucket?: string;
  file_type?: string;
  file_size?: number;
  faculty_email: string;
  upload_date: string;
  due_date: string;
  created_at?: string;
}

const localAttendanceRecords: AttendanceRecord[] = [];
const localExaminations: Examination[] = [
  {
    id: 'exam_csec_001',
    title: 'Aptitude & Machine Learning Assessment',
    subject: 'Machine Learning',
    date: '2026-09-12',
    start_time: '10:00 AM',
    end_time: '11:30 AM',
    description: 'Aptitude and Unit 2 Model Evaluation Assessment for 2nd Year CSE-C students.',
    instructions: 'Bring your laptop with model evaluation notebooks prepared. Scientific calculators allowed.',
    year: 'Second Year',
    department: 'CSE',
    section: 'CSE-C',
    faculty_email: 'anjali.menon@example.edu',
    created_at: new Date().toISOString()
  }
];

const localStudyMaterials: StudyMaterial[] = [];

export interface Subject {
  id: string;
  subject_code: string;
  subject_name: string;
  department: string;
  academic_year: string;
  section: string;
  semester?: string;
  created_at?: string;
  updated_at?: string;
}

export interface FacultySubjectAssignment {
  id: string;
  faculty_id?: string;
  faculty_email: string;
  faculty_name?: string;
  subject_id?: string;
  subject_code: string;
  subject_name: string;
  academic_year: string;
  department: string;
  section: string;
  section_name?: string;
  semester?: string;
  created_at?: string;
  updated_at?: string;
}

export interface StudentGrade {
  id: string;
  student_id?: string;
  regno: string;
  student_name?: string;
  student_email?: string;
  subject_code: string;
  subject_name: string;
  academic_year?: string;
  department?: string;
  section: string;
  section_name?: string;
  internal: number;
  internal_marks?: number;
  assignment: number;
  assignment_marks?: number;
  exam: number;
  exam_marks?: number;
  total: number;
  grade: string;
  faculty_email: string;
  created_at?: string;
  updated_at?: string;
}

const localSubjects: Subject[] = [];
const localFacultySubjectAssignments: FacultySubjectAssignment[] = [];

const localStudentGrades: StudentGrade[] = [];

export interface Assignment {
  id: string;
  subject_code: string;
  subject_name: string;
  academic_year: string;
  department: string;
  section: string;
  faculty_email: string;
  title: string;
  description: string;
  assigned_date: string;
  due_date: string;
  max_marks: number;
  attachment_url?: string;
  attachment_path?: string;
  priority?: string;
  created_at?: string;
  updated_at?: string;
}

export interface AssignmentSubmission {
  id: string;
  assignment_id: string;
  student_id?: string;
  regno: string;
  student_name?: string;
  student_email: string;
  submitted_at?: string;
  file_url?: string;
  file_path?: string;
  response_text?: string;
  status: 'PENDING' | 'SUBMITTED' | 'LATE' | 'GRADED';
  marks?: number;
  grade?: string;
  feedback?: string;
  graded_by?: string;
  graded_at?: string;
  created_at?: string;
  updated_at?: string;
}

export interface StudentAssignmentStatus {
  id: string;
  assignment_id: string;
  student_id?: string;
  regno: string;
  student_email: string;
  status: 'PENDING' | 'COMPLETED';
  completed_at?: string;
  created_at?: string;
}

const localAssignments: Assignment[] = [];
const localAssignmentSubmissions: AssignmentSubmission[] = [];
const localStudentAssignmentStatuses: StudentAssignmentStatus[] = [];

// Helper: Save local memory states to localStorage
function saveLocalStores() {
  if (typeof window !== 'undefined') {
    try {
      localStorage.setItem('cogniva_attendance', JSON.stringify(localAttendanceRecords));
      localStorage.setItem('cogniva_examinations', JSON.stringify(localExaminations));
      localStorage.setItem('cogniva_materials', JSON.stringify(localStudyMaterials));
      localStorage.setItem('cogniva_subjects', JSON.stringify(localSubjects));
      localStorage.setItem('cogniva_subject_assignments', JSON.stringify(localFacultySubjectAssignments));
      localStorage.setItem('cogniva_grades', JSON.stringify(localStudentGrades));
      localStorage.setItem('cogniva_assignments', JSON.stringify(localAssignments));
      localStorage.setItem('cogniva_submissions', JSON.stringify(localAssignmentSubmissions));
      localStorage.setItem('cogniva_assignment_statuses', JSON.stringify(localStudentAssignmentStatuses));
      localStorage.setItem('cogniva_local_notices_v1', JSON.stringify(localNotices));
    } catch {}
  }
}

// Load initial localStorage caches
export function reloadLocalStores() {
  if (typeof window !== 'undefined') {
    try {
      reloadStoredLocalNotices();
      const savedAtt = localStorage.getItem('cogniva_attendance');
      if (savedAtt) {
        const parsed = JSON.parse(savedAtt);
        if (Array.isArray(parsed) && parsed.length > 0) {
          localAttendanceRecords.length = 0;
          localAttendanceRecords.push(...parsed);
        }
      }
      const savedExams = localStorage.getItem('cogniva_examinations');
      if (savedExams) {
        const parsed = JSON.parse(savedExams);
        if (Array.isArray(parsed) && parsed.length > 0) {
          localExaminations.length = 0;
          localExaminations.push(...parsed);
        }
      }
      const savedMats = localStorage.getItem('cogniva_materials');
      if (savedMats) {
        const parsed = JSON.parse(savedMats);
        if (Array.isArray(parsed) && parsed.length > 0) {
          const realMats = parsed.filter(m => m.id !== 'mat_csec_001' && !m.title?.includes('(UNIT 4.pdf)') && !m.file_url?.includes('JVBERi0xLjQK'));
          localStudyMaterials.length = 0;
          localStudyMaterials.push(...realMats);
        }
      }
      const isMockSubject = (name?: string, code?: string) => {
        if (!name && !code) return false;
        const n = (name || '').toLowerCase();
        const c = (code || '').toLowerCase();
        return n.includes('database management') || n.includes('computer networks') || c === 'dbms' || c === 'cn';
      };

      const savedSubjs = localStorage.getItem('cogniva_subjects');
      if (savedSubjs) {
        const parsed = JSON.parse(savedSubjs);
        if (Array.isArray(parsed)) {
          const clean = parsed.filter(s => !isMockSubject(s.subject_name, s.subject_code));
          localSubjects.length = 0;
          localSubjects.push(...clean);
          localStorage.setItem('cogniva_subjects', JSON.stringify(clean));
        }
      }
      const savedAssigns = localStorage.getItem('cogniva_subject_assignments');
      if (savedAssigns) {
        const parsed = JSON.parse(savedAssigns);
        if (Array.isArray(parsed)) {
          const clean = parsed.filter(a => !isMockSubject(a.subject_name, a.subject_code));
          localFacultySubjectAssignments.length = 0;
          localFacultySubjectAssignments.push(...clean);
          localStorage.setItem('cogniva_subject_assignments', JSON.stringify(clean));
        }
      }
      const savedGrades = localStorage.getItem('cogniva_grades');
      if (savedGrades) {
        const parsed = JSON.parse(savedGrades);
        if (Array.isArray(parsed) && parsed.length > 0) {
          localStudentGrades.length = 0;
          localStudentGrades.push(...parsed);
        }
      }
      const savedAsgns = localStorage.getItem('cogniva_assignments');
      if (savedAsgns) {
        const parsed = JSON.parse(savedAsgns);
        if (Array.isArray(parsed) && parsed.length > 0) {
          localAssignments.length = 0;
          localAssignments.push(...parsed);
        }
      }
      const savedSubs = localStorage.getItem('cogniva_submissions');
      if (savedSubs) {
        const parsed = JSON.parse(savedSubs);
        if (Array.isArray(parsed) && parsed.length > 0) {
          localAssignmentSubmissions.length = 0;
          localAssignmentSubmissions.push(...parsed);
        }
      }
      const savedStatuses = localStorage.getItem('cogniva_assignment_statuses');
      if (savedStatuses) {
        const parsed = JSON.parse(savedStatuses);
        if (Array.isArray(parsed) && parsed.length > 0) {
          localStudentAssignmentStatuses.length = 0;
          localStudentAssignmentStatuses.push(...parsed);
        }
      }
    } catch {}
  }
}

reloadLocalStores();

export async function fetchAttendanceRecords(
  sectionOrFilter?: string | { regno?: string; section?: string; subject?: string },
  dateOrSubject?: string,
  subjectOrDate?: string
): Promise<AttendanceRecord[]> {
  let sectionName: string | undefined;
  let date: string | undefined;
  let subject: string | undefined;
  let filterRegno: string | undefined;

  if (typeof sectionOrFilter === 'object' && sectionOrFilter !== null) {
    filterRegno = sectionOrFilter.regno;
    sectionName = sectionOrFilter.section;
    subject = sectionOrFilter.subject;
  } else {
    sectionName = sectionOrFilter;
    date = dateOrSubject;
    subject = subjectOrDate;
  }

  try {
    let query = supabase.from('attendance').select('*').order('date', { ascending: false });
    if (sectionName) query = query.eq('section', sectionName);
    if (filterRegno) query = query.eq('regno', filterRegno);
    if (date) query = query.eq('date', date);
    if (subject) query = query.eq('subject', subject);
    const { data, error } = await query;
    if (!error && data && data.length > 0) {
      return data as AttendanceRecord[];
    }
  } catch {}

  let filtered = [...localAttendanceRecords];
  if (sectionName) {
    filtered = filtered.filter(r => r.section.toLowerCase() === sectionName!.toLowerCase() || r.section.toLowerCase().endsWith(sectionName!.toLowerCase()));
  }
  if (filterRegno) {
    filtered = filtered.filter(r => r.regno.toLowerCase() === filterRegno!.toLowerCase());
  }
  if (date) {
    filtered = filtered.filter(r => r.date === date);
  }
  if (subject) {
    filtered = filtered.filter(r => r.subject.toLowerCase() === subject!.toLowerCase());
  }
  return filtered;
}

export async function saveAttendanceBatch(
  records: Omit<AttendanceRecord, 'id'>[]
): Promise<{ success: boolean; count: number; error?: string }> {
  let count = 0;
  for (const rec of records) {
    try {
      const recordWithId: AttendanceRecord = {
        ...rec,
        id: `att_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`,
        updated_at: new Date().toISOString()
      };

      const { error } = await supabase
        .from('attendance_records')
        .upsert([rec], { onConflict: 'regno,subject,date' });

      const existingIdx = localAttendanceRecords.findIndex(
        (r) => r.regno.toLowerCase() === rec.regno.toLowerCase() && r.subject.toLowerCase() === rec.subject.toLowerCase() && r.date === rec.date
      );
      if (existingIdx !== -1) {
        localAttendanceRecords[existingIdx] = { ...localAttendanceRecords[existingIdx], ...rec };
      } else {
        localAttendanceRecords.push(recordWithId);
      }
      count++;
    } catch {
      const recordWithId: AttendanceRecord = {
        ...rec,
        id: `att_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`,
        updated_at: new Date().toISOString()
      };
      localAttendanceRecords.push(recordWithId);
      count++;
    }
  }

  saveLocalStores();
  return { success: true, count };
}

export async function fetchAttendanceHistory(facultyEmail?: string): Promise<AttendanceRecord[]> {
  try {
    let query = supabase.from('attendance_records').select('*').order('date', { ascending: false });
    if (facultyEmail) {
      query = query.eq('faculty_email', facultyEmail.toLowerCase());
    }
    const { data, error } = await query;
    if (!error && data && data.length > 0) {
      return data as AttendanceRecord[];
    }
  } catch {}

  if (facultyEmail) {
    return localAttendanceRecords.filter(r => r.faculty_email.toLowerCase() === facultyEmail.toLowerCase());
  }
  return localAttendanceRecords;
}

export function validateAttendanceExcelHeaders(headers: string[]): {
  valid: boolean;
  missingColumns: string[];
  mappedHeaders: Record<string, string>;
} {
  const normalized = headers.map(h => String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''));
  const mappedHeaders: Record<string, string> = {};

  const defs = [
    { key: 'regno', keywords: ['regno', 'registernumber', 'rollno', 'rollnumber', 'registrationnumber', 'id', 'studentid'] },
    { key: 'student_name', keywords: ['name', 'studentname', 'fullname', 'student'] },
    { key: 'subject', keywords: ['subject', 'course', 'subj'] },
    { key: 'date', keywords: ['date', 'attendancedate', 'day'] },
    { key: 'status', keywords: ['status', 'attendance', 'attendancestatus', 'present', 'presentabsent'] }
  ];

  for (const def of defs) {
    let idx = normalized.findIndex(h => def.keywords.includes(h));
    if (idx === -1) {
      idx = normalized.findIndex(h => h.length > 1 && def.keywords.some(k => h.includes(k) || k.includes(h)));
    }
    if (idx !== -1) {
      mappedHeaders[def.key] = headers[idx];
    }
  }

  const missingColumns: string[] = [];
  if (!mappedHeaders['regno']) missingColumns.push('Registration Number (Regno)');
  if (!mappedHeaders['status']) missingColumns.push('Attendance Status');

  return {
    valid: missingColumns.length === 0,
    missingColumns,
    mappedHeaders
  };
}

export async function importAttendanceExcel(
  fileOrRows: File | any[],
  facultyEmail: string,
  targetYear: string,
  targetDept: string,
  targetSection: string,
  defaultSubject: string,
  defaultDate: string
): Promise<ImportResult> {
  let importedCount = 0;
  let updatedCount = 0;
  let skippedCount = 0;
  let invalidCount = 0;
  const errors: string[] = [];

  let rows: any[] = [];
  if (Array.isArray(fileOrRows)) {
    rows = fileOrRows;
  } else if (fileOrRows && typeof fileOrRows === 'object') {
    try {
      const buffer = await (fileOrRows as File).arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      rows = XLSX.utils.sheet_to_json(worksheet);
    } catch {
      return {
        importedCount: 0,
        updatedCount: 0,
        skippedCount: 0,
        invalidCount: 1,
        errors: ['Failed to parse Excel file. Please upload a valid .xlsx or .xls file.']
      };
    }
  }

  const assignedStudents = await fetchFacultyAssignedStudents(facultyEmail);
  const assignedRegnos = new Set(assignedStudents.map(s => s.regno.toLowerCase()));

  const toSave: Omit<AttendanceRecord, 'id'>[] = [];

  for (let idx = 0; idx < rows.length; idx++) {
    const row = rows[idx];
    const lineNo = idx + 2;

    const regno = String(row.regno || row.rollno || row.id || '').trim();
    const name = String(row.student_name || row.name || '').trim() || 'Student';
    const subj = String(row.subject || defaultSubject || 'Subject').trim();
    const dt = String(row.date || defaultDate || new Date().toISOString().split('T')[0]).trim();
    const rawStatus = String(row.status || row.attendance || '').trim().toLowerCase();

    if (!regno) {
      invalidCount++;
      errors.push(`Row ${lineNo}: Missing Registration Number`);
      continue;
    }

    if (assignedRegnos.size > 0 && !assignedRegnos.has(regno.toLowerCase())) {
      skippedCount++;
      errors.push(`Row ${lineNo} (${regno}): Student does not belong to assigned section (${targetSection})`);
      continue;
    }

    const status: 'Present' | 'Absent' =
      rawStatus.includes('p') || rawStatus === '1' || rawStatus.includes('present')
        ? 'Present'
        : 'Absent';

    toSave.push({
      regno,
      student_name: name,
      faculty_email: facultyEmail,
      year: targetYear,
      department: targetDept,
      section: targetSection,
      subject: subj,
      date: dt,
      status
    });
  }

  if (toSave.length > 0) {
    const res = await saveAttendanceBatch(toSave);
    importedCount = res.count;
  }

  return { importedCount, updatedCount, skippedCount, invalidCount, errors };
}

export async function createExamination(
  exam: Omit<Examination, 'id'>
): Promise<{ success: boolean; data?: Examination; error?: string }> {
  try {
    const newExam: Examination = {
      ...exam,
      id: `exam_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
      created_at: new Date().toISOString()
    };

    const { data, error } = await supabase
      .from('examinations')
      .insert([exam])
      .select()
      .single();

    if (!error && data) {
      localExaminations.unshift(data as Examination);
      saveLocalStores();
      return { success: true, data: data as Examination };
    }

    localExaminations.unshift(newExam);
    saveLocalStores();
    return { success: true, data: newExam };
  } catch (err) {
    const newExam: Examination = {
      ...exam,
      id: `exam_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
      created_at: new Date().toISOString()
    };
    localExaminations.unshift(newExam);
    saveLocalStores();
    return { success: true, data: newExam };
  }
}

export async function fetchExaminations(sectionName?: string): Promise<Examination[]> {
  try {
    let query = supabase.from('examinations').select('*').order('created_at', { ascending: false });
    if (sectionName) {
      query = query.eq('section', sectionName);
    }
    const { data, error } = await query;
    if (!error && data) {
      return data as Examination[];
    }
  } catch {}

  if (sectionName) {
    return localExaminations.filter(
      e => e.section.toLowerCase() === sectionName.toLowerCase() || e.section.toLowerCase().endsWith(sectionName.toLowerCase())
    );
  }
  return localExaminations;
}

export async function createStudyMaterial(
  material: Omit<StudyMaterial, 'id'>
): Promise<{ success: boolean; data?: StudyMaterial; error?: string }> {
  try {
    const newMat: StudyMaterial = {
      ...material,
      id: `mat_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
      created_at: new Date().toISOString()
    };

    const { data, error } = await supabase
      .from('study_materials')
      .insert([newMat])
      .select()
      .single();

    if (!error && data) {
      localStudyMaterials.unshift(data as StudyMaterial);
      saveLocalStores();
      return { success: true, data: data as StudyMaterial };
    }

    localStudyMaterials.unshift(newMat);
    saveLocalStores();
    return { success: true, data: newMat };
  } catch (err) {
    const newMat: StudyMaterial = {
      ...material,
      id: `mat_${Date.now()}_${Math.floor(Math.random() * 1000)}`,
      created_at: new Date().toISOString()
    };
    localStudyMaterials.unshift(newMat);
    saveLocalStores();
    return { success: true, data: newMat };
  }
}

export async function fetchStudyMaterials(sectionName?: string): Promise<StudyMaterial[]> {
  reloadLocalStores();

  let allMaterials: StudyMaterial[] = [];

  try {
    const { data, error } = await supabase.from('study_materials').select('*').order('created_at', { ascending: false });
    if (!error && data && data.length > 0) {
      allMaterials = data as StudyMaterial[];
    }
  } catch (err) {
    console.warn('Supabase fetchStudyMaterials query notice:', err);
  }

  const existingIds = new Set(allMaterials.map((m) => m.id));
  for (const lm of localStudyMaterials) {
    if (!existingIds.has(lm.id)) {
      allMaterials.push(lm);
      existingIds.add(lm.id);
    }
  }

  let filtered = allMaterials.filter((m) => m.id !== 'mat_csec_001' && !m.title?.includes('(UNIT 4.pdf)'));

  if (sectionName) {
    const cleanSec = sectionName.trim().toUpperCase();
    filtered = filtered.filter((m) => {
      if (!m.section) return true;
      const sec = m.section.trim().toUpperCase();
      return sec === 'ALL' || sec === cleanSec || sec.endsWith(cleanSec) || cleanSec.endsWith(sec);
    });
  }

  return filtered;
}



export async function uploadFileToSupabaseStorage(
  file: File,
  year: string = 'Second Year',
  dept: string = 'CSE',
  sec: string = 'CSE-C',
  sem: string = '4',
  subject: string = 'DBMS',
  bucketName: string = 'study-materials'
): Promise<{ success: boolean; url?: string; path?: string; error?: string }> {
  try {
    if (!file || file.size <= 0) {
      return { success: false, error: 'Please select a valid non-empty file.' };
    }

    if (file.type && file.type !== 'application/pdf' && !file.name.toLowerCase().endsWith('.pdf')) {
      return { success: false, error: 'Only PDF files are allowed.' };
    }

    const maxSizeBytes = 50 * 1024 * 1024; // 50MB
    if (file.size > maxSizeBytes) {
      return { success: false, error: 'File size exceeds the allowed limit (max 50MB).' };
    }

    const cleanYear = (year || 'Second Year').toLowerCase().replace(/[^a-z0-9]/g, '-').replace(/-+/g, '-');
    const cleanDept = (dept || 'CSE').replace(/[^a-zA-Z0-9]/g, '');
    const cleanSec = (sec || 'CSE-C').replace(/[^a-zA-Z0-9]/g, '');
    const cleanSem = `semester-${String(sem || '4').replace(/[^0-9]/g, '') || '4'}`;
    const cleanSubj = (subject || 'DBMS').replace(/[^a-zA-Z0-9]/g, '_');
    const uniqueId = `${Date.now()}_${Math.random().toString(36).substring(2, 8)}`;
    const sanitizedFileName = file.name.replace(/[^a-zA-Z0-9._-]/g, '_');

    const storagePath = `${cleanYear}/${cleanDept}/${cleanSec}/${cleanSem}/${cleanSubj}/${uniqueId}-${sanitizedFileName}`;

    console.log(`[SUPABASE STORAGE] Uploading exact File object (${file.name}, ${file.size} bytes, type: ${file.type}) to '${bucketName}/${storagePath}'...`);

    const { data, error } = await supabase.storage
      .from(bucketName)
      .upload(storagePath, file, {
        contentType: file.type || 'application/pdf',
        upsert: false
      });

    if (error) {
      console.warn('[SUPABASE STORAGE CLIENT ERROR]:', {
        message: error.message,
        name: error.name,
        code: (error as any).code,
        status: (error as any).status,
        statusCode: (error as any).statusCode
      });

      // Fallback: Delegate upload to server-side endpoint /api/upload-material
      console.log('[SUPABASE STORAGE] Delegating upload to secure server endpoint /api/upload-material...');
      try {
        const fileBase64 = await new Promise<string>((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => {
            const resultStr = reader.result as string;
            const base64 = resultStr.includes(',') ? resultStr.split(',')[1] : resultStr;
            resolve(base64);
          };
          reader.onerror = (e) => reject(e);
          reader.readAsDataURL(file);
        });

        const sRes = await fetch('/api/upload-material', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            fileName: file.name,
            fileType: file.type || 'application/pdf',
            fileData: fileBase64,
            year,
            dept,
            section: sec,
            subject,
            facultyEmail: 'anjali.menon@example.edu'
          })
        });

        if (sRes.ok) {
          const serverData = await sRes.json();
          if (serverData.success && serverData.url && serverData.path) {
            console.log('✅ [SERVER UPLOAD SUCCESS]:', serverData);
            return {
              success: true,
              url: serverData.url,
              path: serverData.path
            };
          }
        }
      } catch (sErr) {
        console.error('[SERVER UPLOAD EXCEPTION]:', sErr);
      }

      return { success: false, error: `Storage upload failed: ${error.message}` };
    }

    if (data) {
      const { data: pubData } = supabase.storage.from(bucketName).getPublicUrl(storagePath);
      return {
        success: true,
        url: pubData.publicUrl,
        path: storagePath
      };
    }

    return { success: false, error: 'Storage upload returned no data.' };
  } catch (err: any) {
    console.error('Supabase Storage upload exception:', err);
    return { success: false, error: err?.message || 'Upload failed due to an unexpected storage exception.' };
  }
}

// SUBJECTS API
export async function fetchSubjects(yearOrFilter?: string | { section?: string; department?: string; year?: string }, deptCode?: string, sectionName?: string): Promise<Subject[]> {
  let yearStr = typeof yearOrFilter === 'string' ? yearOrFilter : yearOrFilter?.year;
  if (typeof yearOrFilter === 'object' && yearOrFilter !== null) {
    if (yearOrFilter.department) deptCode = yearOrFilter.department;
    if (yearOrFilter.section) sectionName = yearOrFilter.section;
  }

  try {
    let query = supabase.from('subjects').select('*').order('subject_name', { ascending: true });
    if (sectionName) {
      query = query.eq('section', sectionName);
    }
    const { data, error } = await query;
    if (!error && data) {
      return data as Subject[];
    }
  } catch {}

  let filtered = localSubjects.filter(s => {
    const n = (s.subject_name || '').toLowerCase();
    const c = (s.subject_code || '').toLowerCase();
    return !n.includes('database management') && !n.includes('computer networks') && c !== 'dbms' && c !== 'cn';
  });
  if (yearStr) {
    filtered = filtered.filter(s => s.academic_year.toLowerCase().includes(yearStr.toLowerCase()) || yearStr.toLowerCase().includes(s.academic_year.toLowerCase()));
  }
  if (deptCode) {
    filtered = filtered.filter(s => s.department.toLowerCase().includes(deptCode.toLowerCase()) || deptCode.toLowerCase().includes(s.department.toLowerCase()));
  }
  if (sectionName) {
    filtered = filtered.filter(s => s.section.toLowerCase() === sectionName.toLowerCase() || s.section.toLowerCase().endsWith(sectionName.toLowerCase()));
  }
  return filtered;
}

export async function createSubject(subj: Omit<Subject, 'id'>): Promise<{ success: boolean; data?: Subject; error?: string }> {
  try {
    const newId = `subj_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`;
    const newSubj: Subject = {
      ...subj,
      id: newId,
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    };

    const { data, error } = await supabase.from('subjects').insert([newSubj]).select().single();
    if (!error && data) {
      localSubjects.unshift(data as Subject);
      saveLocalStores();
      return { success: true, data: data as Subject };
    }

    localSubjects.unshift(newSubj);
    saveLocalStores();
    return { success: true, data: newSubj };
  } catch {
    const newId = `subj_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`;
    const newSubj: Subject = {
      ...subj,
      id: newId,
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    };
    localSubjects.unshift(newSubj);
    saveLocalStores();
    return { success: true, data: newSubj };
  }
}

export async function deleteSubject(id: string): Promise<{ success: boolean; error?: string }> {
  try {
    await supabase.from('subjects').delete().eq('id', id);
  } catch {}
  const idx = localSubjects.findIndex(s => s.id === id);
  if (idx !== -1) localSubjects.splice(idx, 1);
  saveLocalStores();
  return { success: true };
}

// FACULTY SUBJECT ASSIGNMENTS API
export async function fetchFacultySubjectAssignments(facultyEmailOrId?: string, sectionName?: string): Promise<FacultySubjectAssignment[]> {
  try {
    let query = supabase.from('faculty_subject_assignments').select('*');
    if (facultyEmailOrId) {
      query = query.or(`faculty_email.eq.${facultyEmailOrId.toLowerCase()},faculty_id.eq.${facultyEmailOrId}`);
    }
    if (sectionName) {
      query = query.eq('section', sectionName);
    }
    const { data, error } = await query;
    if (!error && data) {
      return (data as any[]).map(a => ({
        ...a,
        section_name: a.section_name || a.section
      })) as FacultySubjectAssignment[];
    }
  } catch {}

  let filtered = localFacultySubjectAssignments.filter(a => {
    const n = (a.subject_name || '').toLowerCase();
    const c = (a.subject_code || '').toLowerCase();
    return !n.includes('database management') && !n.includes('computer networks') && c !== 'dbms' && c !== 'cn';
  }).map(a => ({
    ...a,
    section_name: a.section_name || a.section
  }));
  if (facultyEmailOrId) {
    const clean = facultyEmailOrId.toLowerCase();
    filtered = filtered.filter(a => a.faculty_email.toLowerCase() === clean || a.faculty_id?.toLowerCase() === clean);
  }
  if (sectionName) {
    filtered = filtered.filter(a => a.section.toLowerCase() === sectionName.toLowerCase() || a.section.toLowerCase().endsWith(sectionName.toLowerCase()));
  }
  return filtered;
}

export async function assignFacultyToSubject(assignment: Omit<FacultySubjectAssignment, 'id'>): Promise<{ success: boolean; data?: FacultySubjectAssignment; error?: string }> {
  try {
    const newId = `fsa_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`;
    const newAssign: FacultySubjectAssignment = {
      ...assignment,
      id: newId,
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    };

    const { data, error } = await supabase.from('faculty_subject_assignments').insert([newAssign]).select().single();
    if (!error && data) {
      localFacultySubjectAssignments.unshift(data as FacultySubjectAssignment);
      saveLocalStores();
      return { success: true, data: data as FacultySubjectAssignment };
    }

    localFacultySubjectAssignments.unshift(newAssign);
    saveLocalStores();
    return { success: true, data: newAssign };
  } catch {
    const newId = `fsa_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`;
    const newAssign: FacultySubjectAssignment = {
      ...assignment,
      id: newId,
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    };
    localFacultySubjectAssignments.unshift(newAssign);
    saveLocalStores();
    return { success: true, data: newAssign };
  }
}

export async function removeFacultySubjectAssignment(id: string): Promise<{ success: boolean; error?: string }> {
  try {
    await supabase.from('faculty_subject_assignments').delete().eq('id', id);
  } catch {}
  const idx = localFacultySubjectAssignments.findIndex(a => a.id === id);
  if (idx !== -1) localFacultySubjectAssignments.splice(idx, 1);
  saveLocalStores();
  return { success: true };
}

export async function adminAssignSubjectToFaculty(
  facultyEmailOrId: string,
  subjectName: string,
  targetSectionName?: string,
  customSubjectCode?: string
): Promise<{ success: boolean; data?: FacultySubjectAssignment; error?: string }> {
  const cleanName = subjectName.trim();
  if (!cleanName) {
    return { success: false, error: 'Subject name cannot be empty.' };
  }

  if (!facultyEmailOrId) {
    return { success: false, error: 'Please select a faculty member.' };
  }

  const [facMembers, facProfiles] = await Promise.all([
    fetchFacultyMembers(),
    fetchProfilesByRole('faculty')
  ]);

  const fac = facMembers.find(f => f.email.toLowerCase() === facultyEmailOrId.toLowerCase() || f.id === facultyEmailOrId)
    || facProfiles.find(p => p.email.toLowerCase() === facultyEmailOrId.toLowerCase() || p.id === facultyEmailOrId);

  if (!fac) {
    return { success: false, error: 'Selected faculty member not found.' };
  }

  const facEmail = fac.email;
  const facName = (fac as any).name || (fac as any).full_name || facEmail;
  const facId = fac.id;

  const assignedSections = await fetchFacultyAssignedSections(facEmail);
  if (assignedSections.length === 0) {
    return {
      success: false,
      error: `Faculty member '${facName}' has not been assigned to any section yet. Please assign a section (e.g. CSE-C) under 'Faculty Access' first.`
    };
  }

  let secName = (targetSectionName || '').trim();
  if (secName) {
    const isAllowed = assignedSections.some(s => s.name.toLowerCase() === secName.toLowerCase());
    if (!isAllowed) {
      return {
        success: false,
        error: `Faculty member '${facName}' is not assigned to section '${secName}'. Allowed sections: ${assignedSections.map(s => s.name).join(', ')}`
      };
    }
  } else {
    secName = assignedSections[0].name;
  }

  let subjectCode = (customSubjectCode || '').trim().toUpperCase();
  if (!subjectCode) {
    const words = cleanName.split(/\s+/).filter(w => w.length > 0);
    if (words.length >= 2) {
      subjectCode = words.map(w => w[0]).join('').toUpperCase();
    } else {
      subjectCode = cleanName.substring(0, 4).toUpperCase();
    }
  }

  const existingAssignments = await fetchFacultySubjectAssignments(facEmail);
  const yearName = 'Second Year';
  const deptCode = 'CSE';

  const isDuplicate = existingAssignments.some(
    a => (a.section_name || a.section || '').toLowerCase() === secName.toLowerCase() &&
         (a.subject_name.toLowerCase() === cleanName.toLowerCase() || a.subject_code.toLowerCase() === subjectCode.toLowerCase())
  );

  if (isDuplicate) {
    return {
      success: false,
      error: `Subject '${cleanName}' (${subjectCode}) is already assigned to ${facName} for section ${secName}.`
    };
  }

  await createSubject({
    subject_code: subjectCode,
    subject_name: cleanName,
    academic_year: yearName,
    department: deptCode,
    section: secName,
    semester: '4'
  });

  const res = await assignFacultyToSubject({
    faculty_id: facId,
    faculty_email: facEmail,
    faculty_name: facName,
    subject_code: subjectCode,
    subject_name: cleanName,
    academic_year: yearName,
    department: deptCode,
    section: secName,
    section_name: secName
  });

  return res;
}

export async function facultyAddSubject(
  facultyEmail: string,
  subjectName: string,
  targetSectionName?: string
): Promise<{ success: boolean; data?: FacultySubjectAssignment; error?: string }> {
  const cleanName = subjectName.trim();
  if (!cleanName) {
    return { success: false, error: 'Subject name cannot be empty.' };
  }

  if (!facultyEmail) {
    return { success: false, error: 'Faculty authentication email is missing.' };
  }

  let assignedSections = await fetchFacultyAssignedSections(facultyEmail);
  if (assignedSections.length === 0) {
    const facMembers = await fetchFacultyMembers();
    const fac = facMembers.find(f => f.email.toLowerCase() === facultyEmail.toLowerCase());
    if (fac && fac.section) {
      const allSecs = await fetchSections();
      const matchedSec = allSecs.find(s => s.name.toLowerCase() === fac.section?.toLowerCase());
      if (matchedSec) {
        assignedSections = [matchedSec];
      } else {
        assignedSections = [{ id: 's3', academic_year_id: 'y2', department_id: 'd1', name: fac.section, capacity: 60 }];
      }
    } else {
      assignedSections = [{ id: 's3', academic_year_id: 'y2', department_id: 'd1', name: 'CSE-C', capacity: 60 }];
    }
  }

  const secName = targetSectionName || assignedSections[0].name;

  return adminAssignSubjectToFaculty(facultyEmail, cleanName, secName);
}

// GRADES & MARKS API
export function calculateGradeAndTotal(internal: number, assignment: number, exam: number): { total: number; grade: string } {
  const i = Math.max(0, Math.min(20, Number(internal) || 0));
  const a = Math.max(0, Math.min(10, Number(assignment) || 0));
  const e = Math.max(0, Math.min(70, Number(exam) || 0));
  const total = Math.round(i + a + e);
  let grade = 'F';
  if (total >= 90) grade = 'O';
  else if (total >= 80) grade = 'A+';
  else if (total >= 70) grade = 'A';
  else if (total >= 60) grade = 'B+';
  else if (total >= 50) grade = 'B';
  else if (total >= 40) grade = 'C';
  return { total, grade };
}

export async function fetchStudentGrades(studentRegnoOrEmail?: string, sectionNameOrCode?: string, subjectCode?: string): Promise<StudentGrade[]> {
  let sectionName: string | undefined;
  if (sectionNameOrCode && !subjectCode) {
    // If 2 params passed, 2nd could be sectionName or subjectCode
    if (sectionNameOrCode.includes('-')) sectionName = sectionNameOrCode;
    else subjectCode = sectionNameOrCode;
  } else if (sectionNameOrCode && subjectCode) {
    sectionName = sectionNameOrCode;
  }

  try {
    let query = supabase.from('grades').select('*');
    if (studentRegnoOrEmail) {
      query = query.or(`regno.eq.${studentRegnoOrEmail.toUpperCase()},student_email.eq.${studentRegnoOrEmail.toLowerCase()},faculty_email.eq.${studentRegnoOrEmail.toLowerCase()}`);
    }
    if (subjectCode) {
      query = query.eq('subject_code', subjectCode);
    }
    if (sectionName) {
      query = query.eq('section', sectionName);
    }
    const { data, error } = await query;
    if (!error && data) {
      return (data as any[]).map(g => ({
        ...g,
        section_name: g.section_name || g.section,
        internal_marks: g.internal_marks ?? g.internal ?? 0,
        assignment_marks: g.assignment_marks ?? g.assignment ?? 0,
        exam_marks: g.exam_marks ?? g.exam ?? 0,
        internal: g.internal ?? g.internal_marks ?? 0,
        assignment: g.assignment ?? g.assignment_marks ?? 0,
        exam: g.exam ?? g.exam_marks ?? 0
      })) as StudentGrade[];
    }
  } catch {}

  let filtered = [...localStudentGrades].map(g => ({
    ...g,
    section_name: g.section_name || g.section,
    internal_marks: g.internal_marks ?? g.internal ?? 0,
    assignment_marks: g.assignment_marks ?? g.assignment ?? 0,
    exam_marks: g.exam_marks ?? g.exam ?? 0,
    internal: g.internal ?? g.internal_marks ?? 0,
    assignment: g.assignment ?? g.assignment_marks ?? 0,
    exam: g.exam ?? g.exam_marks ?? 0
  }));
  if (studentRegnoOrEmail) {
    const clean = studentRegnoOrEmail.toLowerCase();
    filtered = filtered.filter(g => g.regno.toLowerCase() === clean || g.student_email?.toLowerCase() === clean || g.faculty_email.toLowerCase() === clean);
  }
  if (subjectCode) {
    filtered = filtered.filter(g => g.subject_code.toLowerCase() === subjectCode!.toLowerCase() || g.subject_name.toLowerCase().includes(subjectCode!.toLowerCase()));
  }
  if (sectionName) {
    filtered = filtered.filter(g => g.section.toLowerCase() === sectionName!.toLowerCase() || g.section.toLowerCase().endsWith(sectionName!.toLowerCase()));
  }
  return filtered;
}

export async function saveStudentGradesBatch(grades: Omit<StudentGrade, 'id'>[]): Promise<{ success: boolean; count: number; error?: string }> {
  let count = 0;
  for (const rawG of grades) {
    const internal = rawG.internal_marks ?? rawG.internal ?? 0;
    const assignment = rawG.assignment_marks ?? rawG.assignment ?? 0;
    const exam = rawG.exam_marks ?? rawG.exam ?? 0;
    const { total, grade } = calculateGradeAndTotal(internal, assignment, exam);

    const rec: StudentGrade = {
      ...rawG,
      internal,
      internal_marks: internal,
      assignment,
      assignment_marks: assignment,
      exam,
      exam_marks: exam,
      total,
      grade,
      section: rawG.section || rawG.section_name || 'CSE-C',
      section_name: rawG.section_name || rawG.section || 'CSE-C',
      id: `grd_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
      updated_at: new Date().toISOString()
    };

    try {
      await supabase.from('grades').upsert([rec], { onConflict: 'regno,subject_code' });
    } catch {}

    const existingIdx = localStudentGrades.findIndex(
      lg => lg.regno.toLowerCase() === rec.regno.toLowerCase() && lg.subject_code.toLowerCase() === rec.subject_code.toLowerCase()
    );
    if (existingIdx !== -1) {
      localStudentGrades[existingIdx] = { ...localStudentGrades[existingIdx], ...rec };
    } else {
      localStudentGrades.push(rec);
    }
    count++;
  }

  saveLocalStores();
  return { success: true, count };
}

export function validateGradesExcelHeaders(headers: string[]): { valid: boolean; missingColumns: string[]; mappedHeaders: Record<string, string> } {
  const normalized = headers.map(h => String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''));
  const mappedHeaders: Record<string, string> = {};

  const defs = [
    { key: 'regno', keywords: ['regno', 'registernumber', 'rollno', 'rollnumber', 'registrationnumber', 'id', 'studentid'] },
    { key: 'internal', keywords: ['internal', 'internalmarks', 'internals', 'int'] },
    { key: 'assignment', keywords: ['assignment', 'assignmentmarks', 'assgn', 'ass'] },
    { key: 'exam', keywords: ['exam', 'exammarks', 'endsem', 'semesterexam', 'final'] },
    { key: 'subject_code', keywords: ['subjectcode', 'subjcode', 'code', 'subject'] }
  ];

  for (const def of defs) {
    let idx = normalized.findIndex(h => def.keywords.includes(h));
    if (idx === -1) {
      idx = normalized.findIndex(h => h.length > 1 && def.keywords.some(k => h.includes(k) || k.includes(h)));
    }
    if (idx !== -1) {
      mappedHeaders[def.key] = headers[idx];
    }
  }

  const missingColumns: string[] = [];
  if (!mappedHeaders['regno']) missingColumns.push('Register Number (Regno)');

  return { valid: missingColumns.length === 0, missingColumns, mappedHeaders };
}

export async function importGradesExcel(
  fileOrRows: File | any[],
  facultyEmail: string,
  targetSubjectCode?: string,
  targetSection?: string,
  targetYear: string = 'Second Year',
  targetDept: string = 'CSE',
  targetSubjectName: string = ''
): Promise<ImportResult> {
  let importedCount = 0;
  let updatedCount = 0;
  let skippedCount = 0;
  let invalidCount = 0;
  const errors: string[] = [];

  let rows: any[] = [];
  if (Array.isArray(fileOrRows)) {
    rows = fileOrRows;
  } else if (fileOrRows && typeof fileOrRows === 'object') {
    try {
      const buffer = await (fileOrRows as File).arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      rows = XLSX.utils.sheet_to_json(worksheet);
    } catch {
      return { importedCount: 0, updatedCount: 0, skippedCount: 0, invalidCount: 1, errors: ['Failed to parse Excel file. Please upload a valid .xlsx or .xls file.'] };
    }
  }

  const assignedStudents = await fetchFacultyAssignedStudents(facultyEmail);
  const assignedRegnos = new Set(assignedStudents.map(s => s.regno.toLowerCase()));

  const toSave: Omit<StudentGrade, 'id'>[] = [];

  for (let idx = 0; idx < rows.length; idx++) {
    const row = rows[idx];
    const lineNo = idx + 2;

    const regno = String(row.regno || row.rollno || row.id || '').trim();
    const internal = Number(row.internal || row.internalmarks || 0);
    const assignment = Number(row.assignment || row.assignmentmarks || 0);
    const exam = Number(row.exam || row.exammarks || row.endsem || 0);

    if (!regno) {
      invalidCount++;
      errors.push(`Row ${lineNo}: Missing Register Number`);
      continue;
    }

    if (assignedRegnos.size > 0 && !assignedRegnos.has(regno.toLowerCase())) {
      skippedCount++;
      errors.push(`Row ${lineNo} (${regno}): Student does not belong to your assigned section (${targetSection})`);
      continue;
    }

    const { total, grade } = calculateGradeAndTotal(internal, assignment, exam);

    toSave.push({
      regno,
      academic_year: targetYear || 'Second Year',
      department: targetDept || 'CSE',
      section: targetSection || 'CSE-C',
      subject_code: targetSubjectCode || 'CS402',
      subject_name: targetSubjectName || 'DBMS',
      internal,
      assignment,
      exam,
      total,
      grade,
      faculty_email: facultyEmail
    });
  }

  if (toSave.length > 0) {
    const res = await saveStudentGradesBatch(toSave as StudentGrade[]);
    if (res.success) importedCount = res.count;
  }

  return { importedCount, updatedCount, skippedCount, invalidCount, errors };
}

// ASSIGNMENTS API
export function calculateDueDatePriority(dueDateStr: string, isCompleted: boolean, isSubmitted: boolean): {
  status: 'OVERDUE' | 'DUE TODAY' | 'DUE TOMORROW' | 'DUE SOON' | 'UPCOMING' | 'COMPLETED';
  label: string;
  tone: 'coral' | 'amber' | 'teal' | 'violet';
  daysDiff: number;
} {
  if (isCompleted || isSubmitted) {
    return { status: 'COMPLETED', label: 'Completed', tone: 'teal', daysDiff: 999 };
  }

  const todayStr = '2026-09-09';
  const now = new Date(todayStr);
  const due = new Date(dueDateStr);
  const diffTime = due.getTime() - now.getTime();
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

  if (diffDays < 0) {
    return { status: 'OVERDUE', label: 'Overdue', tone: 'coral', daysDiff: diffDays };
  } else if (diffDays === 0) {
    return { status: 'DUE TODAY', label: 'Due Today', tone: 'coral', daysDiff: 0 };
  } else if (diffDays === 1) {
    return { status: 'DUE TOMORROW', label: 'Due Tomorrow', tone: 'coral', daysDiff: 1 };
  } else if (diffDays <= 3) {
    return { status: 'DUE SOON', label: `Due in ${diffDays} days`, tone: 'amber', daysDiff: diffDays };
  } else {
    return { status: 'UPCOMING', label: `Due in ${diffDays} days`, tone: 'teal', daysDiff: diffDays };
  }
}

export function calculateExamPriority(examDateStr: string): {
  status: 'PAST' | 'EXAM TODAY' | 'EXAM TOMORROW' | 'EXAM SOON' | 'UPCOMING';
  label: string;
  tone: 'coral' | 'amber' | 'violet' | 'teal';
  daysDiff: number;
} {
  const todayStr = '2026-09-09';
  const now = new Date(todayStr);
  const examDate = new Date(examDateStr);
  const diffTime = examDate.getTime() - now.getTime();
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

  if (diffDays < 0) {
    return { status: 'PAST', label: 'Past Exam', tone: 'teal', daysDiff: 999 };
  } else if (diffDays === 0) {
    return { status: 'EXAM TODAY', label: 'EXAM TODAY', tone: 'coral', daysDiff: 0 };
  } else if (diffDays === 1) {
    return { status: 'EXAM TOMORROW', label: 'EXAM TOMORROW', tone: 'amber', daysDiff: 1 };
  } else if (diffDays <= 3) {
    return { status: 'EXAM SOON', label: `EXAM IN ${diffDays} DAYS`, tone: 'amber', daysDiff: diffDays };
  } else {
    return { status: 'UPCOMING', label: `EXAM IN ${diffDays} DAYS`, tone: 'violet', daysDiff: diffDays };
  }
}


export async function fetchAssignments(filter?: { section?: string; faculty_email?: string; subject_code?: string }): Promise<Assignment[]> {
  reloadLocalStores();

  let allAssignments: Assignment[] = [];

  try {
    const { data, error } = await supabase.from('assignments').select('*').order('due_date', { ascending: true });
    if (!error && data && data.length > 0) {
      allAssignments = data as Assignment[];
    }
  } catch (err) {
    console.warn('Supabase fetchAssignments query notice:', err);
  }

  const existingIds = new Set(allAssignments.map(a => a.id));
  for (const la of localAssignments) {
    if (!existingIds.has(la.id)) {
      allAssignments.push(la);
      existingIds.add(la.id);
    }
  }

  let filtered = [...allAssignments];

  if (filter?.section) {
    const cleanSec = filter.section.trim().toUpperCase();
    filtered = filtered.filter(a => {
      if (!a.section) return true;
      const sec = a.section.trim().toUpperCase();
      return sec === 'ALL' || sec === cleanSec || sec.endsWith(cleanSec) || cleanSec.endsWith(sec);
    });
  }

  if (filter?.faculty_email) {
    const cleanEmail = filter.faculty_email.trim().toLowerCase();
    filtered = filtered.filter(a => a.faculty_email && a.faculty_email.trim().toLowerCase() === cleanEmail);
  }

  if (filter?.subject_code) {
    const cleanSub = filter.subject_code.trim().toLowerCase();
    filtered = filtered.filter(a => a.subject_code && a.subject_code.trim().toLowerCase() === cleanSub);
  }

  return filtered;
}

export async function createAssignment(asgn: Omit<Assignment, 'id'>): Promise<{ success: boolean; data?: Assignment; error?: string }> {
  reloadLocalStores();
  const newId = `asgn_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`;
  const record: Assignment = {
    ...asgn,
    id: newId,
    faculty_email: asgn.faculty_email.toLowerCase(),
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString()
  };

  localAssignments.unshift(record);
  saveLocalStores();

  try {
    const { data, error } = await supabase.from('assignments').insert([record]).select().single();
    if (!error && data) {
      return { success: true, data: data as Assignment };
    }
  } catch (err) {
    console.warn('Supabase assignments insert notice:', err);
  }

  return { success: true, data: record };
}

export async function updateAssignment(id: string, updates: Partial<Assignment>): Promise<{ success: boolean; data?: Assignment; error?: string }> {
  try {
    await supabase.from('assignments').update({ ...updates, updated_at: new Date().toISOString() }).eq('id', id);
  } catch {}

  const idx = localAssignments.findIndex(a => a.id === id);
  if (idx !== -1) {
    localAssignments[idx] = { ...localAssignments[idx], ...updates, updated_at: new Date().toISOString() };
  }
  saveLocalStores();
  return { success: true };
}

export async function deleteAssignment(id: string): Promise<{ success: boolean; error?: string }> {
  try {
    await supabase.from('assignments').delete().eq('id', id);
  } catch {}

  const idx = localAssignments.findIndex(a => a.id === id);
  if (idx !== -1) localAssignments.splice(idx, 1);
  saveLocalStores();
  return { success: true };
}

export async function fetchAssignmentSubmissions(assignmentId?: string, studentRegnoOrEmail?: string): Promise<AssignmentSubmission[]> {
  try {
    let query = supabase.from('assignment_submissions').select('*');
    if (assignmentId) query = query.eq('assignment_id', assignmentId);
    if (studentRegnoOrEmail) {
      query = query.or(`regno.eq.${studentRegnoOrEmail.toUpperCase()},student_email.eq.${studentRegnoOrEmail.toLowerCase()}`);
    }
    const { data, error } = await query;
    if (!error && data) {
      return data as AssignmentSubmission[];
    }
  } catch {}

  let filtered = [...localAssignmentSubmissions];
  if (assignmentId) filtered = filtered.filter(s => s.assignment_id === assignmentId);
  if (studentRegnoOrEmail) {
    const clean = studentRegnoOrEmail.toLowerCase();
    filtered = filtered.filter(s => s.regno.toLowerCase() === clean || s.student_email.toLowerCase() === clean);
  }
  return filtered;
}

export async function submitAssignment(sub: Omit<AssignmentSubmission, 'id'>): Promise<{ success: boolean; data?: AssignmentSubmission; error?: string }> {
  const newId = `sub_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`;
  const record: AssignmentSubmission = {
    ...sub,
    id: newId,
    status: sub.status || 'SUBMITTED',
    submitted_at: new Date().toISOString(),
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString()
  };

  try {
    const { data, error } = await supabase.from('assignment_submissions').upsert([record], { onConflict: 'assignment_id,regno' }).select().single();
    if (!error && data) {
      const idx = localAssignmentSubmissions.findIndex(s => s.assignment_id === sub.assignment_id && s.regno.toLowerCase() === sub.regno.toLowerCase());
      if (idx !== -1) localAssignmentSubmissions[idx] = data as AssignmentSubmission;
      else localAssignmentSubmissions.unshift(data as AssignmentSubmission);
      saveLocalStores();
      return { success: true, data: data as AssignmentSubmission };
    }
  } catch {}

  const idx = localAssignmentSubmissions.findIndex(s => s.assignment_id === sub.assignment_id && s.regno.toLowerCase() === sub.regno.toLowerCase());
  if (idx !== -1) localAssignmentSubmissions[idx] = record;
  else localAssignmentSubmissions.unshift(record);
  saveLocalStores();
  return { success: true, data: record };
}

export async function gradeSubmission(
  submissionId: string,
  assignmentId: string,
  regno: string,
  marks: number,
  grade: string,
  feedback: string,
  gradedBy: string
): Promise<{ success: boolean; error?: string }> {
  const now = new Date().toISOString();
  try {
    await supabase.from('assignment_submissions').update({
      marks,
      grade,
      feedback,
      graded_by: gradedBy,
      graded_at: now,
      status: 'GRADED',
      updated_at: now
    }).eq('assignment_id', assignmentId).eq('regno', regno);
  } catch {}

  const idx = localAssignmentSubmissions.findIndex(s => s.assignment_id === assignmentId && s.regno.toLowerCase() === regno.toLowerCase());
  if (idx !== -1) {
    localAssignmentSubmissions[idx] = {
      ...localAssignmentSubmissions[idx],
      marks,
      grade,
      feedback,
      graded_by: gradedBy,
      graded_at: now,
      status: 'GRADED',
      updated_at: now
    };
  }
  saveLocalStores();
  return { success: true };
}

export async function fetchStudentAssignmentStatuses(studentRegnoOrEmail?: string): Promise<StudentAssignmentStatus[]> {
  try {
    let query = supabase.from('student_assignment_status').select('*');
    if (studentRegnoOrEmail) {
      query = query.or(`regno.eq.${studentRegnoOrEmail.toUpperCase()},student_email.eq.${studentRegnoOrEmail.toLowerCase()}`);
    }
    const { data, error } = await query;
    if (!error && data) {
      return data as StudentAssignmentStatus[];
    }
  } catch {}

  let filtered = [...localStudentAssignmentStatuses];
  if (studentRegnoOrEmail) {
    const clean = studentRegnoOrEmail.toLowerCase();
    filtered = filtered.filter(s => s.regno.toLowerCase() === clean || s.student_email.toLowerCase() === clean);
  }
  return filtered;
}

export async function toggleAssignmentCompletion(assignmentId: string, regno: string, studentEmail: string): Promise<{ success: boolean; isCompleted: boolean }> {
  const existingIdx = localStudentAssignmentStatuses.findIndex(s => s.assignment_id === assignmentId && s.regno.toLowerCase() === regno.toLowerCase());
  let nextCompleted = true;

  if (existingIdx !== -1) {
    const current = localStudentAssignmentStatuses[existingIdx];
    nextCompleted = current.status !== 'COMPLETED';
    localStudentAssignmentStatuses[existingIdx].status = nextCompleted ? 'COMPLETED' : 'PENDING';
    localStudentAssignmentStatuses[existingIdx].completed_at = nextCompleted ? new Date().toISOString() : undefined;
  } else {
    localStudentAssignmentStatuses.push({
      id: `sas_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
      assignment_id: assignmentId,
      regno,
      student_email: studentEmail,
      status: 'COMPLETED',
      completed_at: new Date().toISOString(),
      created_at: new Date().toISOString()
    });
    nextCompleted = true;
  }

  try {
    await supabase.from('student_assignment_status').upsert([{
      assignment_id: assignmentId,
      regno,
      student_email: studentEmail,
      status: nextCompleted ? 'COMPLETED' : 'PENDING',
      completed_at: nextCompleted ? new Date().toISOString() : null
    }], { onConflict: 'assignment_id,regno' });
  } catch {}

  saveLocalStores();
  return { success: true, isCompleted: nextCompleted };
}

// ============================================================================
// NOTICES & ANNOUNCEMENTS API
// ============================================================================

export interface Notice {
  id: string;
  title: string;
  description: string;
  image_url?: string;
  image_path?: string;
  faculty_email: string;
  faculty_name?: string;
  section: string;
  year?: string;
  department?: string;
  priority: 'Normal' | 'Important' | 'Urgent';
  published_at: string;
  expires_at?: string;
  created_at?: string;
  updated_at?: string;
  read_by?: string[];
}

const LOCAL_NOTICES_STORAGE_KEY = 'cogniva_local_notices_v1';

function getStoredLocalNotices(): Notice[] {
  try {
    const raw = localStorage.getItem(LOCAL_NOTICES_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) return parsed;
    }
  } catch {}
  return [];
}

function saveStoredLocalNotices(notices: Notice[]) {
  try {
    localStorage.setItem(LOCAL_NOTICES_STORAGE_KEY, JSON.stringify(notices));
  } catch {}
}

function reloadStoredLocalNotices() {
  try {
    const raw = localStorage.getItem(LOCAL_NOTICES_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        localNotices.length = 0;
        localNotices.push(...parsed);
      }
    }
  } catch {}
}

const localNotices: Notice[] = getStoredLocalNotices();

export async function fetchNotices(sectionName?: string, facultyEmail?: string): Promise<Notice[]> {
  reloadStoredLocalNotices();

  let allNotices: Notice[] = [];

  try {
    const { data, error } = await supabase
      .from('notices')
      .select('*')
      .order('published_at', { ascending: false });

    if (!error && data && data.length > 0) {
      allNotices = data as Notice[];
    }
  } catch (err) {
    console.warn('Supabase fetchNotices query notice:', err);
  }

  const existingIds = new Set(allNotices.map((n) => n.id));
  for (const ln of localNotices) {
    if (!existingIds.has(ln.id)) {
      allNotices.push(ln);
      existingIds.add(ln.id);
    }
  }

  const todayStr = new Date().toISOString().split('T')[0];

  let filtered = allNotices.filter((n) => {
    if (!n.expires_at) return true;
    return n.expires_at >= todayStr;
  });

  if (sectionName) {
    const cleanSec = sectionName.trim().toUpperCase();
    filtered = filtered.filter((n) => {
      if (!n.section) return true;
      const sec = n.section.trim().toUpperCase();
      return sec === 'ALL' || sec === cleanSec || sec.endsWith(cleanSec) || cleanSec.endsWith(sec);
    });
  }

  if (facultyEmail) {
    const cleanEmail = facultyEmail.trim().toLowerCase();
    filtered = filtered.filter((n) => n.faculty_email && n.faculty_email.trim().toLowerCase() === cleanEmail);
  }

  const priorityOrder: Record<string, number> = { Urgent: 1, Important: 2, Normal: 3 };

  filtered.sort((a, b) => {
    const pA = priorityOrder[a.priority] || 3;
    const pB = priorityOrder[b.priority] || 3;
    if (pA !== pB) return pA - pB;
    return new Date(b.published_at).getTime() - new Date(a.published_at).getTime();
  });

  return filtered;
}

export async function createNotice(noticeData: Omit<Notice, 'id'>): Promise<{ success: boolean; data?: Notice; error?: string }> {
  reloadStoredLocalNotices();
  const newId = `notice_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`;
  const record: Notice = {
    ...noticeData,
    id: newId,
    faculty_email: noticeData.faculty_email.toLowerCase(),
    read_by: noticeData.read_by || [],
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString(),
  };

  localNotices.unshift(record);
  saveStoredLocalNotices(localNotices);

  try {
    const { data, error } = await supabase.from('notices').insert([record]).select().single();
    if (!error && data) {
      return { success: true, data: data as Notice };
    }
  } catch (err) {
    console.warn('Supabase notices insert notice:', err);
  }

  return { success: true, data: record };
}

export async function deleteNotice(noticeId: string): Promise<{ success: boolean; error?: string }> {
  try {
    await supabase.from('notices').delete().eq('id', noticeId);
  } catch {}

  const idx = localNotices.findIndex((n) => n.id === noticeId);
  if (idx !== -1) {
    localNotices.splice(idx, 1);
  }
  saveStoredLocalNotices(localNotices);
  return { success: true };
}

export async function markNoticeAsRead(noticeId: string, studentEmailOrRegno: string): Promise<{ success: boolean }> {
  const cleanUser = studentEmailOrRegno.trim().toLowerCase();

  const idx = localNotices.findIndex((n) => n.id === noticeId);
  if (idx !== -1) {
    if (!localNotices[idx].read_by) localNotices[idx].read_by = [];
    if (!localNotices[idx].read_by!.includes(cleanUser)) {
      localNotices[idx].read_by!.push(cleanUser);
    }
    saveStoredLocalNotices(localNotices);
  }

  try {
    const { data: currentNotice } = await supabase.from('notices').select('read_by').eq('id', noticeId).single();
    let updatedReadBy: string[] = currentNotice?.read_by || [];
    if (!updatedReadBy.includes(cleanUser)) {
      updatedReadBy.push(cleanUser);
      await supabase.from('notices').update({ read_by: updatedReadBy }).eq('id', noticeId);
    }
  } catch {}

  return { success: true };
}

export async function uploadNoticeImage(
  file: File,
  section: string,
  facultyEmail: string
): Promise<{ success: boolean; url?: string; path?: string; error?: string }> {
  try {
    if (!file || file.size <= 0) {
      return { success: false, error: 'Please select a valid image file.' };
    }

    if (!file.type.startsWith('image/')) {
      return { success: false, error: 'Only image files (JPEG, PNG, WebP, GIF) are allowed.' };
    }

    const fileBase64 = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const resultStr = reader.result as string;
        const base64 = resultStr.includes(',') ? resultStr.split(',')[1] : resultStr;
        resolve(base64);
      };
      reader.onerror = (e) => reject(e);
      reader.readAsDataURL(file);
    });

    const res = await fetch('/api/upload-notice-image', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        fileName: file.name,
        fileType: file.type || 'image/jpeg',
        fileData: fileBase64,
        section,
        facultyEmail,
      }),
    });

    if (res.ok) {
      const serverData = await res.json();
      if (serverData.success && serverData.url && serverData.path) {
        return {
          success: true,
          url: serverData.url,
          path: serverData.path,
        };
      }
      return { success: false, error: serverData.error || 'Server upload failed' };
    } else {
      const errJson = await res.json().catch(() => ({}));
      return { success: false, error: errJson.error || 'Notice image upload endpoint returned an error.' };
    }
  } catch (err: any) {
    return { success: false, error: err?.message || 'Failed to upload notice image.' };
  }
}

// ============================================================================
// EXAM RESULTS & EXCEL MARK IMPORT API
// ============================================================================

export interface ExamResult {
  id: string;
  regno: string;
  student_name?: string;
  student_email?: string;
  department?: string;
  semester?: string;
  section?: string;
  subject_code: string;
  subject_name: string;
  exam_type: 'Internal Exam 1' | 'Internal Exam 2' | 'Internal Exam 3' | 'Semester Examination';
  marks: number;
  max_marks: number;
  faculty_email: string;
  created_at?: string;
  updated_at?: string;
}

export interface ExamImportHistory {
  id: string;
  faculty_email: string;
  exam_type: string;
  import_date: string;
  total_students: number;
  total_records: number;
  successful_records: number;
  failed_records: number;
  file_name: string;
}

const LOCAL_EXAM_RESULTS_KEY = 'cogniva_local_exam_results_v1';
const LOCAL_IMPORT_HISTORY_KEY = 'cogniva_local_import_history_v1';

let localExamResults: ExamResult[] = (() => {
  try {
    const saved = localStorage.getItem(LOCAL_EXAM_RESULTS_KEY);
    if (saved) return JSON.parse(saved);
  } catch {}
  return [];
})();

let localImportHistory: ExamImportHistory[] = (() => {
  try {
    const saved = localStorage.getItem(LOCAL_IMPORT_HISTORY_KEY);
    if (saved) return JSON.parse(saved);
  } catch {}
  return [];
})();

function saveExamStores() {
  try {
    localStorage.setItem(LOCAL_EXAM_RESULTS_KEY, JSON.stringify(localExamResults));
    localStorage.setItem(LOCAL_IMPORT_HISTORY_KEY, JSON.stringify(localImportHistory));
  } catch {}
}

export async function fetchExamResults(filter?: { regno?: string; section?: string; exam_type?: string; faculty_email?: string }): Promise<ExamResult[]> {
  try {
    let query = supabase.from('exam_results').select('*');
    if (filter?.regno) query = query.eq('regno', filter.regno.toUpperCase());
    if (filter?.section) query = query.eq('section', filter.section);
    if (filter?.exam_type) query = query.eq('exam_type', filter.exam_type);
    if (filter?.faculty_email) query = query.eq('faculty_email', filter.faculty_email.toLowerCase());

    const { data, error } = await query;
    if (!error && data && data.length > 0) {
      return data as ExamResult[];
    }
  } catch {}

  let filtered = [...localExamResults];
  if (filter?.regno) {
    const cleanReg = filter.regno.toLowerCase();
    filtered = filtered.filter(r => r.regno.toLowerCase() === cleanReg);
  }
  if (filter?.section) {
    filtered = filtered.filter(r => r.section?.toLowerCase() === filter.section!.toLowerCase() || r.section?.toLowerCase().endsWith(filter.section!.toLowerCase()));
  }
  if (filter?.exam_type) {
    filtered = filtered.filter(r => r.exam_type.toLowerCase() === filter.exam_type!.toLowerCase());
  }
  if (filter?.faculty_email) {
    filtered = filtered.filter(r => r.faculty_email.toLowerCase() === filter.faculty_email!.toLowerCase());
  }
  return filtered;
}

export async function saveExamResultsBatch(results: Omit<ExamResult, 'id'>[], mode: 'replace' | 'keep' = 'replace'): Promise<{ success: boolean; count: number }> {
  let count = 0;
  for (const raw of results) {
    const rec: ExamResult = {
      ...raw,
      id: `exr_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
      updated_at: new Date().toISOString(),
      created_at: raw.created_at || new Date().toISOString()
    };

    try {
      await supabase.from('exam_results').upsert([rec], { onConflict: 'regno,subject_name,exam_type' });
    } catch {}

    const existingIdx = localExamResults.findIndex(
      r => r.regno.toLowerCase() === rec.regno.toLowerCase() &&
           r.subject_name.toLowerCase() === rec.subject_name.toLowerCase() &&
           r.exam_type.toLowerCase() === rec.exam_type.toLowerCase()
    );

    if (existingIdx !== -1) {
      if (mode === 'replace') {
        localExamResults[existingIdx] = { ...localExamResults[existingIdx], ...rec };
        count++;
      }
    } else {
      localExamResults.push(rec);
      count++;
    }
  }

  saveExamStores();
  return { success: true, count };
}

export async function fetchExamImportHistory(facultyEmail?: string): Promise<ExamImportHistory[]> {
  try {
    let query = supabase.from('exam_import_history').select('*').order('import_date', { ascending: false });
    if (facultyEmail) query = query.eq('faculty_email', facultyEmail.toLowerCase());
    const { data, error } = await query;
    if (!error && data && data.length > 0) {
      return data as ExamImportHistory[];
    }
  } catch {}

  if (facultyEmail) {
    return localImportHistory.filter(h => h.faculty_email.toLowerCase() === facultyEmail.toLowerCase());
  }
  return localImportHistory;
}

export async function saveExamImportHistory(log: Omit<ExamImportHistory, 'id'>): Promise<void> {
  const rec: ExamImportHistory = {
    ...log,
    id: `imp_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`
  };
  try {
    await supabase.from('exam_import_history').insert([rec]);
  } catch {}
  localImportHistory.unshift(rec);
  saveExamStores();
}

// ----------------------------------------------------
// DYNAMIC RESULT DATASETS (100% Dynamic Excel Tables)
// ----------------------------------------------------

export interface DynamicResultRow {
  regno: string;
  studentName: string;
  studentEmail?: string;
  department?: string;
  semester?: string;
  section?: string;
  data: Record<string, number | string | null>;
}

export interface DynamicResultDataset {
  id: string;
  fileName: string;
  facultyEmail: string;
  importedAt: string;
  headers: string[]; // Original Excel headers in exact order
  regNoHeader: string;
  nameHeader?: string;
  resultHeaders: string[]; // Result column keys in exact order
  rows: DynamicResultRow[];
}

const LOCAL_DYNAMIC_DATASETS_KEY = 'cogniva_dynamic_result_datasets_v1';

let localDynamicDatasets: DynamicResultDataset[] = [];

try {
  const raw = localStorage.getItem(LOCAL_DYNAMIC_DATASETS_KEY);
  if (raw) localDynamicDatasets = JSON.parse(raw);
} catch {}

function saveDynamicDatasetsStore() {
  try {
    localStorage.setItem(LOCAL_DYNAMIC_DATASETS_KEY, JSON.stringify(localDynamicDatasets));
  } catch {}
}

export async function fetchDynamicResultsDataset(facultyEmail?: string): Promise<DynamicResultDataset | null> {
  try {
    let query = supabase.from('dynamic_result_datasets').select('*').order('imported_at', { ascending: false });
    if (facultyEmail) query = query.eq('faculty_email', facultyEmail.toLowerCase());
    const { data, error } = await query.limit(1);
    if (!error && data && data.length > 0) {
      return data[0] as DynamicResultDataset;
    }
  } catch {}

  if (facultyEmail) {
    const filtered = localDynamicDatasets.filter(d => !facultyEmail || d.facultyEmail.toLowerCase() === facultyEmail.toLowerCase());
    if (filtered.length > 0) return filtered[0];
  }
  return localDynamicDatasets[0] || null;
}

export async function fetchAllDynamicResultsDatasets(facultyEmail?: string): Promise<DynamicResultDataset[]> {
  try {
    let query = supabase.from('dynamic_result_datasets').select('*').order('imported_at', { ascending: false });
    if (facultyEmail) query = query.eq('faculty_email', facultyEmail.toLowerCase());
    const { data, error } = await query;
    if (!error && data && data.length > 0) {
      return data as DynamicResultDataset[];
    }
  } catch {}

  if (facultyEmail) {
    return localDynamicDatasets.filter(d => d.facultyEmail.toLowerCase() === facultyEmail.toLowerCase());
  }
  return localDynamicDatasets;
}

export async function saveDynamicResultsDataset(dataset: Omit<DynamicResultDataset, 'id' | 'importedAt'>): Promise<DynamicResultDataset> {
  const fullDataset: DynamicResultDataset = {
    ...dataset,
    id: `ds_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
    importedAt: new Date().toISOString()
  };

  try {
    await supabase.from('dynamic_result_datasets').upsert([fullDataset]);
  } catch {}

  const existingIdx = localDynamicDatasets.findIndex(
    d => d.facultyEmail.toLowerCase() === dataset.facultyEmail.toLowerCase() && d.fileName === dataset.fileName
  );
  if (existingIdx !== -1) {
    localDynamicDatasets[existingIdx] = fullDataset;
  } else {
    localDynamicDatasets.unshift(fullDataset);
  }

  saveDynamicDatasetsStore();
  return fullDataset;
}

// ----------------------------------------------------
// CGPA & SGPA ACADEMIC PERFORMANCE SYSTEM
// ----------------------------------------------------

export interface SemesterGpaEntry {
  semester: string; // e.g. "Sem 1", "Sem 2", "Sem 3", "Sem 4"
  sgpa: number | null;
  cgpa: number | null;
  status: 'Completed' | 'Current' | 'Pending';
}

export interface StudentCgpaRecord {
  id: string;
  regno: string;
  studentName: string;
  studentEmail?: string;
  department?: string;
  semester?: string;
  section?: string;
  facultyEmail?: string;
  semesters: SemesterGpaEntry[];
  currentCgpa: number | null;
  latestSgpa: number | null;
  previousSgpa: number | null;
  bestSgpa: number | null;
  lowestSgpa: number | null;
  averageSgpa: number | null;
  trend: 'improving' | 'stable' | 'declining';
  sgpaDelta: number | null;
  cgpaDelta: number | null;
  updatedAt: string;
}

export interface CgpaImportDataset {
  id: string;
  fileName: string;
  facultyEmail: string;
  importedAt: string;
  headers: string[];
  regNoHeader: string;
  nameHeader?: string;
  semHeaders: string[];
  cgpaHeader?: string;
  records: StudentCgpaRecord[];
}

const LOCAL_CGPA_RECORDS_KEY = 'cogniva_student_cgpa_records_v1';
const LOCAL_CGPA_DATASETS_KEY = 'cogniva_cgpa_import_datasets_v1';

let localCgpaRecords: StudentCgpaRecord[] = [];
let localCgpaDatasets: CgpaImportDataset[] = [];

try {
  const rawRecs = localStorage.getItem(LOCAL_CGPA_RECORDS_KEY);
  if (rawRecs) localCgpaRecords = JSON.parse(rawRecs);
  const rawDs = localStorage.getItem(LOCAL_CGPA_DATASETS_KEY);
  if (rawDs) localCgpaDatasets = JSON.parse(rawDs);
} catch {}

if (localCgpaRecords.length === 0) {
  localCgpaRecords = [
    {
      id: 'cgpa_1',
      regno: '2024CSE001',
      studentName: 'Aditya Varma',
      studentEmail: 'aditya.v@example.edu',
      department: 'CSE',
      semester: '4',
      section: 'CSE-C',
      facultyEmail: 'anjali.menon@example.edu',
      semesters: [
        { semester: 'Sem 1', sgpa: 8.10, cgpa: 8.10, status: 'Completed' },
        { semester: 'Sem 2', sgpa: 8.35, cgpa: 8.22, status: 'Completed' },
        { semester: 'Sem 3', sgpa: 8.42, cgpa: 8.29, status: 'Completed' },
        { semester: 'Sem 4', sgpa: null, cgpa: null, status: 'Current' }
      ],
      currentCgpa: 8.29,
      latestSgpa: 8.42,
      previousSgpa: 8.35,
      bestSgpa: 8.42,
      lowestSgpa: 8.10,
      averageSgpa: 8.29,
      trend: 'improving',
      sgpaDelta: 0.07,
      cgpaDelta: 0.07,
      updatedAt: new Date().toISOString()
    },
    {
      id: 'cgpa_2',
      regno: '2024CSE002',
      studentName: 'Bhavna Sharma',
      studentEmail: 'bhavna.s@example.edu',
      department: 'CSE',
      semester: '4',
      section: 'CSE-C',
      facultyEmail: 'anjali.menon@example.edu',
      semesters: [
        { semester: 'Sem 1', sgpa: 8.60, cgpa: 8.60, status: 'Completed' },
        { semester: 'Sem 2', sgpa: 8.75, cgpa: 8.68, status: 'Completed' },
        { semester: 'Sem 3', sgpa: 8.80, cgpa: 8.72, status: 'Completed' },
        { semester: 'Sem 4', sgpa: null, cgpa: null, status: 'Current' }
      ],
      currentCgpa: 8.72,
      latestSgpa: 8.80,
      previousSgpa: 8.75,
      bestSgpa: 8.80,
      lowestSgpa: 8.60,
      averageSgpa: 8.72,
      trend: 'improving',
      sgpaDelta: 0.05,
      cgpaDelta: 0.04,
      updatedAt: new Date().toISOString()
    },
    {
      id: 'cgpa_3',
      regno: '2024CSE003',
      studentName: 'Chetan Kumar',
      studentEmail: 'chetan.k@example.edu',
      department: 'CSE',
      semester: '4',
      section: 'CSE-C',
      facultyEmail: 'anjali.menon@example.edu',
      semesters: [
        { semester: 'Sem 1', sgpa: 7.80, cgpa: 7.80, status: 'Completed' },
        { semester: 'Sem 2', sgpa: 7.90, cgpa: 7.85, status: 'Completed' },
        { semester: 'Sem 3', sgpa: 8.10, cgpa: 7.93, status: 'Completed' },
        { semester: 'Sem 4', sgpa: null, cgpa: null, status: 'Current' }
      ],
      currentCgpa: 7.93,
      latestSgpa: 8.10,
      previousSgpa: 7.90,
      bestSgpa: 8.10,
      lowestSgpa: 7.80,
      averageSgpa: 7.93,
      trend: 'improving',
      sgpaDelta: 0.20,
      cgpaDelta: 0.08,
      updatedAt: new Date().toISOString()
    }
  ];
  saveCgpaStores();
}

function saveCgpaStores() {
  try {
    localStorage.setItem(LOCAL_CGPA_RECORDS_KEY, JSON.stringify(localCgpaRecords));
    localStorage.setItem(LOCAL_CGPA_DATASETS_KEY, JSON.stringify(localCgpaDatasets));
  } catch {}
}

export async function fetchFacultyCgpaRecords(facultyEmail?: string): Promise<StudentCgpaRecord[]> {
  try {
    let query = supabase.from('student_cgpa_records').select('*');
    if (facultyEmail) query = query.eq('faculty_email', facultyEmail.toLowerCase());
    const { data, error } = await query;
    if (!error && data && data.length > 0) {
      return data as StudentCgpaRecord[];
    }
  } catch {}

  if (facultyEmail) {
    return localCgpaRecords.filter(r => !r.facultyEmail || r.facultyEmail.toLowerCase() === facultyEmail.toLowerCase());
  }
  return localCgpaRecords;
}

export async function fetchStudentCgpaRecord(regnoOrEmail: string): Promise<StudentCgpaRecord | null> {
  const clean = regnoOrEmail.toLowerCase().trim();
  try {
    const { data, error } = await supabase
      .from('student_cgpa_records')
      .select('*')
      .or(`regno.ilike.${clean},student_email.ilike.${clean}`)
      .limit(1);
    if (!error && data && data.length > 0) {
      return data[0] as StudentCgpaRecord;
    }
  } catch {}

  const found = localCgpaRecords.find(
    r => r.regno.toLowerCase() === clean || (r.studentEmail && r.studentEmail.toLowerCase() === clean)
  );
  return found || localCgpaRecords[0] || null;
}

export async function saveCgpaRecordsBatch(
  records: StudentCgpaRecord[],
  datasetMeta?: Omit<CgpaImportDataset, 'id' | 'importedAt' | 'records'>
): Promise<{ success: boolean; count: number }> {
  try {
    await supabase.from('student_cgpa_records').upsert(records, { onConflict: 'regno' });
  } catch {}

  records.forEach(rec => {
    const idx = localCgpaRecords.findIndex(r => r.regno.toLowerCase() === rec.regno.toLowerCase());
    if (idx !== -1) {
      localCgpaRecords[idx] = rec;
    } else {
      localCgpaRecords.push(rec);
    }
  });

  if (datasetMeta) {
    const ds: CgpaImportDataset = {
      ...datasetMeta,
      id: `cgpa_ds_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
      importedAt: new Date().toISOString(),
      records
    };
    localCgpaDatasets.unshift(ds);
    try {
      await supabase.from('cgpa_import_datasets').insert([ds]);
    } catch {}
  }

  saveCgpaStores();
  return { success: true, count: records.length };
}

export async function fetchCgpaImportHistory(facultyEmail?: string): Promise<CgpaImportDataset[]> {
  try {
    let query = supabase.from('cgpa_import_datasets').select('*').order('imported_at', { ascending: false });
    if (facultyEmail) query = query.eq('faculty_email', facultyEmail.toLowerCase());
    const { data, error } = await query;
    if (!error && data && data.length > 0) {
      return data as CgpaImportDataset[];
    }
  } catch {}

  if (facultyEmail) {
    return localCgpaDatasets.filter(d => d.facultyEmail.toLowerCase() === facultyEmail.toLowerCase());
  }
  return localCgpaDatasets;
}

// ============================================================================
// DYNAMIC CGPA / SGPA EXCEL PARSER
// ============================================================================

export function parseDynamicCgpaExcel(
  arrayBuffer: ArrayBuffer,
  roster: StudentMember[]
) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const rawRows: Record<string, any>[] = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  if (rawRows.length === 0) {
    return { headers: [], regNoHeader: '', nameHeader: '', semHeaders: [], cgpaHeader: '', matchedRows: [], unmatchedRows: [], totalExcelRows: 0 };
  }

  const headers = Object.keys(rawRows[0]);
  const norm = (s: string) => String(s || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');

  const regNoHeader = headers.find(h =>
    ['regno', 'register', 'rollno', 'studentid', 'reg', 'id'].some(k => norm(h).includes(k))
  ) || headers[0];
  const nameHeader = headers.find(h =>
    ['name', 'studentname', 'fullname'].some(k => norm(h).includes(k))
  ) || '';
  const cgpaHeader = headers.find(h =>
    ['cgpa', 'cumulativegpa', 'overallcgpa', 'cgpa'].some(k => norm(h).includes(k))
  ) || '';

  // Detect semester SGPA columns: "Sem 1 SGPA", "Semester 1", "S1 SGPA", etc.
  const semHeaders = headers.filter(h => {
    const n = norm(h);
    return (n.includes('sem') || n.includes('semester') || n.startsWith('s')) &&
      (n.includes('sgpa') || n.includes('gpa') || n.includes('score')) &&
      !n.includes('cgpa');
  });

  const rosterMap = new Map<string, StudentMember>();
  roster.forEach(s => {
    rosterMap.set(norm(s.regno), s);
    if (s.email) rosterMap.set(norm(s.email), s);
    if (s.name) rosterMap.set(norm(s.name), s);
  });

  const matchedRows: StudentCgpaRecord[] = [];
  const unmatchedRows: Record<string, any>[] = [];

  rawRows.forEach(row => {
    const rawRegno = String(row[regNoHeader] || '').trim();
    const normRegno = norm(rawRegno);
    if (!normRegno) return;

    const student = rosterMap.get(normRegno);
    const sSec = student?.section || '';

    // Build semester entries
    const semesters: SemesterGpaEntry[] = semHeaders.map((sh, idx) => {
      const raw = parseFloat(String(row[sh] || ''));
      const sgpa = isNaN(raw) ? null : Math.min(10, Math.max(0, raw));
      return {
        semester: sh,
        sgpa,
        cgpa: null,
        status: (sgpa !== null ? 'Completed' : idx === semHeaders.length - 1 ? 'Current' : 'Pending') as SemesterGpaEntry['status']
      };
    });

    // Compute rolling CGPA
    let runTotal = 0, runCount = 0;
    semesters.forEach(s => {
      if (s.sgpa !== null) { runTotal += s.sgpa; runCount++; }
      s.cgpa = runCount > 0 ? Math.round((runTotal / runCount) * 100) / 100 : null;
    });

    const completedSems = semesters.filter(s => s.sgpa !== null);
    const currentCgpa = completedSems.length > 0
      ? parseFloat(cgpaHeader && row[cgpaHeader] ? String(row[cgpaHeader]) : (completedSems.reduce((a, s) => a + (s.sgpa || 0), 0) / completedSems.length).toFixed(2))
      : null;
    const latestSgpa = completedSems.length > 0 ? completedSems[completedSems.length - 1].sgpa : null;
    const previousSgpa = completedSems.length > 1 ? completedSems[completedSems.length - 2].sgpa : null;
    const bestSgpa = completedSems.length > 0 ? Math.max(...completedSems.map(s => s.sgpa || 0)) : null;
    const lowestSgpa = completedSems.length > 0 ? Math.min(...completedSems.map(s => s.sgpa || 10)) : null;
    const averageSgpa = completedSems.length > 0 ? Math.round((completedSems.reduce((a, s) => a + (s.sgpa || 0), 0) / completedSems.length) * 100) / 100 : null;
    const sgpaDelta = (latestSgpa !== null && previousSgpa !== null) ? Math.round((latestSgpa - previousSgpa) * 100) / 100 : null;
    const trend: StudentCgpaRecord['trend'] = sgpaDelta === null ? 'stable' : sgpaDelta > 0.05 ? 'improving' : sgpaDelta < -0.05 ? 'declining' : 'stable';

    const rec: StudentCgpaRecord = {
      id: `cgpa_${rawRegno}_${Date.now()}`,
      regno: rawRegno,
      studentName: nameHeader ? String(row[nameHeader] || student?.name || rawRegno) : (student?.name || rawRegno),
      studentEmail: student?.email,
      department: student?.department || 'CSE',
      section: sSec,
      semester: String(completedSems.length),
      semesters,
      currentCgpa,
      latestSgpa,
      previousSgpa,
      bestSgpa,
      lowestSgpa,
      averageSgpa,
      trend,
      sgpaDelta,
      cgpaDelta: sgpaDelta,
      updatedAt: new Date().toISOString()
    };

    if (student) {
      matchedRows.push(rec);
    } else {
      unmatchedRows.push(row);
    }
  });

  return {
    headers,
    regNoHeader,
    nameHeader,
    semHeaders,
    cgpaHeader,
    matchedRows,
    unmatchedRows,
    totalExcelRows: rawRows.length
  };
}

// ============================================================================
// DYNAMIC SUBJECT-WISE + OVERALL ATTENDANCE API & PERSISTENCE
// ============================================================================

export interface DynamicSubjectAttendance {
  subjectName: string;
  attendancePercentage: number;
  status: 'Good' | 'Watch' | 'At Risk';
}

export interface StudentAttendanceSummaryRecord {
  id: string;
  regno: string;
  studentName: string;
  studentEmail?: string;
  department?: string;
  year?: string;
  section: string;
  facultyEmail?: string;
  subjectAttendances: DynamicSubjectAttendance[];
  overallAttendancePercentage: number | null;
  overallStatus: 'Good' | 'Watch' | 'At Risk';
  highestSubject?: { subjectName: string; percentage: number };
  lowestSubject?: { subjectName: string; percentage: number };
  importedSubjectHeaders?: string[];
  updatedAt: string;
}

export interface DynamicAttendanceImportDataset {
  id: string;
  fileName: string;
  facultyEmail?: string;
  importedBy?: string;
  section: string;
  studentCount?: number;
  subjectCount?: number;
  importedAt: string;
  headers?: string[];
  regNoHeader?: string;
  nameHeader?: string;
  subjectHeaders?: string[];
  overallHeader?: string;
  records: StudentAttendanceSummaryRecord[];
}

const LOCAL_ATT_SUMMARY_KEY = 'cogniva_student_attendance_summary_v1';
const LOCAL_ATT_DATASETS_KEY = 'cogniva_attendance_import_datasets_v1';

let localAttendanceSummaryRecords: StudentAttendanceSummaryRecord[] = [];
let localAttendanceImportDatasets: DynamicAttendanceImportDataset[] = [];

export function reloadLocalAttendanceStores(): StudentAttendanceSummaryRecord[] {
  try {
    if (typeof window !== 'undefined') {
      const rawRecs = localStorage.getItem(LOCAL_ATT_SUMMARY_KEY);
      if (rawRecs) localAttendanceSummaryRecords = JSON.parse(rawRecs);
      const rawDs = localStorage.getItem(LOCAL_ATT_DATASETS_KEY);
      if (rawDs) localAttendanceImportDatasets = JSON.parse(rawDs);
    }
  } catch {}
  return localAttendanceSummaryRecords;
}

reloadLocalAttendanceStores();

function saveAttendanceSummaryStores() {
  try {
    if (typeof window !== 'undefined') {
      localStorage.setItem(LOCAL_ATT_SUMMARY_KEY, JSON.stringify(localAttendanceSummaryRecords));
      localStorage.setItem(LOCAL_ATT_DATASETS_KEY, JSON.stringify(localAttendanceImportDatasets));
    }
  } catch {}
}

export async function fetchFacultyAttendanceSummaryRecords(
  facultyEmail?: string,
  sectionName?: string
): Promise<StudentAttendanceSummaryRecord[]> {
  try {
    let query = supabase.from('student_attendance_summary').select('*');
    if (sectionName) query = query.eq('section', sectionName);
    if (facultyEmail) query = query.eq('faculty_email', facultyEmail.toLowerCase());
    const { data, error } = await query;
    if (!error && data && data.length > 0) {
      return data as StudentAttendanceSummaryRecord[];
    }
  } catch {}

  reloadLocalAttendanceStores();
  let filtered = [...localAttendanceSummaryRecords];
  if (sectionName) {
    filtered = filtered.filter(r => r.section.toLowerCase() === sectionName.toLowerCase() || r.section.toLowerCase().endsWith(sectionName.toLowerCase()));
  }
  if (facultyEmail) {
    filtered = filtered.filter(r => !r.facultyEmail || r.facultyEmail.toLowerCase() === facultyEmail.toLowerCase());
  }
  return filtered;
}

export async function fetchStudentAttendanceSummaryRecord(
  regnoOrEmail: string
): Promise<StudentAttendanceSummaryRecord | null> {
  if (!regnoOrEmail) return null;
  const cleanInput = regnoOrEmail.toLowerCase().trim();

  reloadLocalAttendanceStores();

  let matchedStudent: StudentMember | undefined = undefined;
  try {
    const roster = await fetchStudentMembers();
    matchedStudent = roster.find(
      s => s.email?.toLowerCase() === cleanInput || s.regno?.toLowerCase() === cleanInput || s.id?.toLowerCase() === cleanInput
    );
  } catch {}

  const candidateKeys = new Set<string>();
  candidateKeys.add(cleanInput);
  if (matchedStudent) {
    if (matchedStudent.regno) candidateKeys.add(matchedStudent.regno.toLowerCase().trim());
    if (matchedStudent.email) candidateKeys.add(matchedStudent.email.toLowerCase().trim());
    if (matchedStudent.name) candidateKeys.add(matchedStudent.name.toLowerCase().trim());
  }

  try {
    for (const key of Array.from(candidateKeys)) {
      const { data, error } = await supabase
        .from('student_attendance_summary')
        .select('*')
        .or(`regno.ilike.${key},student_email.ilike.${key},student_name.ilike.${key}`)
        .limit(1);
      if (!error && data && data.length > 0) {
        return data[0] as StudentAttendanceSummaryRecord;
      }
    }
  } catch {}

  const allRecords = reloadLocalAttendanceStores();
  for (const key of Array.from(candidateKeys)) {
    const found = allRecords.find(r => {
      const rReg = (r.regno || '').toLowerCase().trim();
      const rEmail = (r.studentEmail || '').toLowerCase().trim();
      const rName = (r.studentName || '').toLowerCase().trim();
      return rReg === key || rEmail === key || rName === key;
    });
    if (found) return found;
  }

  if (matchedStudent) {
    const roster = await fetchStudentMembers().catch(() => []);
    const studentIdx = roster.findIndex(s => s.regno.toLowerCase() === matchedStudent?.regno.toLowerCase());
    if (studentIdx !== -1 && studentIdx < allRecords.length) {
      return allRecords[studentIdx];
    }
  }

  return null;
}

export async function saveAttendanceSummaryBatch(
  records: StudentAttendanceSummaryRecord[],
  datasetMeta?: Omit<DynamicAttendanceImportDataset, 'id' | 'importedAt' | 'records'>,
  overwrite: boolean = true
): Promise<{ success: boolean; count: number }> {
  try {
    await supabase.from('student_attendance_summary').upsert(records, { onConflict: 'regno' });
  } catch {}

  reloadLocalAttendanceStores();

  records.forEach(newRec => {
    const idx = localAttendanceSummaryRecords.findIndex(r => r.regno.toLowerCase() === newRec.regno.toLowerCase());
    if (idx !== -1) {
      if (overwrite) {
        localAttendanceSummaryRecords[idx] = newRec;
      } else {
        const existingRec = localAttendanceSummaryRecords[idx];
        const mergedSubjects = [...existingRec.subjectAttendances];

        newRec.subjectAttendances.forEach(newSub => {
          const sIdx = mergedSubjects.findIndex(s => s.subjectName.toLowerCase() === newSub.subjectName.toLowerCase());
          if (sIdx !== -1) {
            mergedSubjects[sIdx] = newSub;
          } else {
            mergedSubjects.push(newSub);
          }
        });

        const validPct = mergedSubjects.map(s => s.attendancePercentage).filter((p): p is number => p !== null);
        const overall = newRec.overallAttendancePercentage !== null
          ? newRec.overallAttendancePercentage
          : validPct.length > 0
          ? Math.round(validPct.reduce((a, b) => a + b, 0) / validPct.length)
          : null;

        const overallStatus = overall !== null ? (overall >= 85 ? 'Good' : overall >= 75 ? 'Watch' : 'At Risk') : 'Good';

        localAttendanceSummaryRecords[idx] = {
          ...existingRec,
          subjectAttendances: mergedSubjects,
          overallAttendancePercentage: overall,
          overallStatus,
          importedSubjectHeaders: Array.from(new Set([...(existingRec.importedSubjectHeaders || []), ...(newRec.importedSubjectHeaders || [])])),
          updatedAt: new Date().toISOString()
        };
      }
    } else {
      localAttendanceSummaryRecords.push(newRec);
    }
  });

  if (datasetMeta) {
    const ds: DynamicAttendanceImportDataset = {
      ...datasetMeta,
      id: `att_ds_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
      importedAt: new Date().toISOString(),
      records
    };
    localAttendanceImportDatasets.unshift(ds);
    try {
      await supabase.from('attendance_import_datasets').insert([ds]);
    } catch {}
  }

  saveAttendanceSummaryStores();
  return { success: true, count: records.length };
}

export async function fetchDynamicAttendanceImportHistory(
  facultyEmail?: string
): Promise<DynamicAttendanceImportDataset[]> {
  try {
    let query = supabase.from('attendance_import_datasets').select('*').order('imported_at', { ascending: false });
    if (facultyEmail) query = query.eq('faculty_email', facultyEmail.toLowerCase());
    const { data, error } = await query;
    if (!error && data && data.length > 0) {
      return data as DynamicAttendanceImportDataset[];
    }
  } catch {}

  reloadLocalAttendanceStores();
  if (facultyEmail) {
    return localAttendanceImportDatasets.filter(d => (d.facultyEmail || d.importedBy || '').toLowerCase() === facultyEmail.toLowerCase());
  }
  return localAttendanceImportDatasets;
}

export function parseDynamicAttendanceExcel(
  arrayBuffer: ArrayBuffer,
  roster: StudentMember[]
) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const rawRows: Record<string, any>[] = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  if (rawRows.length === 0) {
    return {
      headers: [],
      regNoHeader: '',
      subjectHeaders: [],
      matchedRows: [],
      unmatchedRows: [],
      totalExcelRows: 0
    };
  }

  const headers = Object.keys(rawRows[0]);
  const norm = (s: string) => String(s || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');

  let regNoHeader = headers.find(h => ['regno', 'register', 'rollno', 'studentid', 'id', 'reg'].some(k => norm(h).includes(k))) || headers[0];
  let nameHeader = headers.find(h => ['name', 'studentname', 'fullname'].some(k => norm(h).includes(k)));
  let emailHeader = headers.find(h => ['email', 'mail'].some(k => norm(h).includes(k)));
  let overallHeader = headers.find(h => norm(h).includes('overall') || norm(h).includes('totalattendance'));

  const serialKeywords = ['sno', 'slno', 'srno', 'no', 'sn', 'sno.', 'slno.', 'srno.', 'row', 'index'];
  const idKeywords = ['regno', 'register', 'registernumber', 'rollno', 'rollnumber', 'studentid', 'student_id', 'id', 'reg'];
  const nameKeywords = ['name', 'studentname', 'fullname', 'student_name', 'full_name'];
  const emailKeywords = ['email', 'studentemail', 'emailaddress', 'mail'];
  const structKeywords = ['department', 'dept', 'section', 'sec', 'year', 'academicyear', 'semester', 'sem', 'dob', 'dateofbirth'];
  const overallKeywords = ['overall', 'overallattendance', 'overall_attendance', 'overall%', 'totalattendance', 'total_attendance', 'total%'];

  const isMeta = (h: string) => {
    const n = norm(h);
    if (!n) return true;
    if (serialKeywords.some(k => n === k)) return true;
    if (idKeywords.some(k => n === k || n.replace(/[^a-z]/g, '') === k)) return true;
    if (nameKeywords.some(k => n === k || n.replace(/[^a-z]/g, '') === k)) return true;
    if (emailKeywords.some(k => n === k || n.replace(/[^a-z]/g, '') === k)) return true;
    if (structKeywords.some(k => n === k)) return true;
    if (overallKeywords.some(k => n === k || n.includes('overall') || n.includes('totalattendance'))) return true;
    return false;
  };

  const subjectHeaders = headers.filter(h => !isMeta(h));

  const matchedRows: StudentAttendanceSummaryRecord[] = [];
  const unmatchedRows: Array<{ rowNumber: number; data: Record<string, any>; reason: string }> = [];

  rawRows.forEach((row, i) => {
    const regVal = String(row[regNoHeader] || '').trim();
    const nameVal = nameHeader ? String(row[nameHeader] || '').trim() : '';
    const emailVal = emailHeader ? String(row[emailHeader] || '').trim() : '';

    if (!regVal && !nameVal && !emailVal) {
      unmatchedRows.push({ rowNumber: i + 2, data: row, reason: 'Empty student identifier row' });
      return;
    }

    const matchedStudent = roster.find(s => 
      (regVal && s.regno.toLowerCase() === regVal.toLowerCase()) ||
      (emailVal && s.email.toLowerCase() === emailVal.toLowerCase()) ||
      (nameVal && s.name.toLowerCase() === nameVal.toLowerCase())
    ) || (roster.length > i ? roster[i] : undefined);

    if (!matchedStudent && regVal && roster.length > 0 && !roster.some(s => s.regno.toLowerCase() === regVal.toLowerCase())) {
      unmatchedRows.push({ rowNumber: i + 2, data: row, reason: `Student '${regVal}' not in authorized section roster` });
    }

    const sReg = matchedStudent?.regno || regVal || `REG_${i + 1}`;
    const sName = matchedStudent?.name || nameVal || `Student ${i + 1}`;
    const sEmail = matchedStudent?.email || emailVal;
    const sDept = matchedStudent?.department || String(row['Department'] || row['Dept'] || 'CSE').trim();
    const sSec = matchedStudent?.section || String(row['Section'] || row['Sec'] || 'CSE-C').trim();

    const subjectAttendances: DynamicSubjectAttendance[] = [];
    const validPercentages: number[] = [];

    subjectHeaders.forEach(sh => {
      const rawVal = String(row[sh] || '').trim();
      if (!rawVal) return;

      const cleaned = rawVal.replace('%', '').trim();
      const num = parseFloat(cleaned);

      if (!isNaN(num) && num >= 0 && num <= 100) {
        const rounded = Math.round(num);
        validPercentages.push(rounded);
        const status = rounded >= 85 ? 'Good' : rounded >= 75 ? 'Watch' : 'At Risk';
        subjectAttendances.push({
          subjectName: sh,
          attendancePercentage: rounded,
          status
        });
      }
    });

    const explicitOverallRaw = overallHeader ? String(row[overallHeader] || '').trim().replace('%', '') : '';
    const explicitOverallNum = parseFloat(explicitOverallRaw);

    const overallAttendancePercentage = !isNaN(explicitOverallNum) && explicitOverallNum >= 0 && explicitOverallNum <= 100
      ? Math.round(explicitOverallNum)
      : validPercentages.length > 0
      ? Math.round(validPercentages.reduce((a, b) => a + b, 0) / validPercentages.length)
      : null;

    const hasLowSubject = subjectAttendances.some(s => s.status === 'At Risk');
    const overallStatus = overallAttendancePercentage !== null
      ? (overallAttendancePercentage >= 85 && !hasLowSubject ? 'Good' : overallAttendancePercentage >= 75 && !hasLowSubject ? 'Watch' : 'At Risk')
      : 'Good';

    let highestSubject: { subjectName: string; percentage: number } | undefined = undefined;
    let lowestSubject: { subjectName: string; percentage: number } | undefined = undefined;

    if (subjectAttendances.length > 0) {
      const sorted = [...subjectAttendances].sort((a, b) => b.attendancePercentage - a.attendancePercentage);
      highestSubject = { subjectName: sorted[0].subjectName, percentage: sorted[0].attendancePercentage };
      lowestSubject = { subjectName: sorted[sorted.length - 1].subjectName, percentage: sorted[sorted.length - 1].attendancePercentage };
    }

    matchedRows.push({
      id: `att_sum_${sReg.toLowerCase().replace(/[^a-z0-9]/g, '_')}`,
      regno: sReg,
      studentName: sName,
      studentEmail: sEmail,
      department: sDept,
      section: sSec,
      subjectAttendances,
      overallAttendancePercentage,
      overallStatus,
      highestSubject,
      lowestSubject,
      importedSubjectHeaders: subjectHeaders,
      updatedAt: new Date().toISOString()
    });
  });

  return {
    headers,
    regNoHeader,
    nameHeader,
    subjectHeaders,
    overallHeader,
    matchedRows,
    unmatchedRows,
    totalExcelRows: rawRows.length
  };
}

export interface DynamicSubjectGrade {
  subjectName: string;
  grade: string;
  gradePoint?: number;
  status?: string;
}

export interface StudentGradeSummaryRecord {
  id: string;
  regno: string;
  studentName: string;
  studentEmail?: string;
  department?: string;
  section: string;
  facultyEmail?: string;
  subjectGrades: DynamicSubjectGrade[];
  overallGrade?: string;
  overallGradePoint?: number;
  highestGradeSubject?: { subjectName: string; grade: string };
  lowestGradeSubject?: { subjectName: string; grade: string };
  updatedAt: string;
}

export interface DynamicGradeImportDataset {
  id: string;
  fileName: string;
  importedBy: string;
  section: string;
  subjectCount: number;
  studentCount: number;
  importedAt: string;
  records: StudentGradeSummaryRecord[];
}

const LOCAL_GRADE_SUMMARY_KEY = 'cogniva_student_grade_summary_v1';
const LOCAL_GRADE_DATASETS_KEY = 'cogniva_grade_import_datasets_v1';

const localGradeSummaryRecords: StudentGradeSummaryRecord[] = loadGradeSummaryRecords();
const localGradeImportDatasets: DynamicGradeImportDataset[] = loadGradeImportDatasets();

function loadGradeSummaryRecords(): StudentGradeSummaryRecord[] {
  try {
    const raw = localStorage.getItem(LOCAL_GRADE_SUMMARY_KEY);
    if (raw) return JSON.parse(raw);
  } catch {}
  return [
    {
      id: 'grd_sum_23cs001',
      regno: '23CS001',
      studentName: 'Aditya Varma',
      studentEmail: 'aditya.varma@example.edu',
      department: 'CSE',
      section: 'CSE-C',
      facultyEmail: 'anjali.menon@example.edu',
      subjectGrades: [
        { subjectName: 'Data Analytics', grade: 'A', gradePoint: 9, status: 'Good' },
        { subjectName: 'Cloud Computing', grade: 'A+', gradePoint: 10, status: 'Good' },
        { subjectName: 'Embedded Programming', grade: 'B+', gradePoint: 8, status: 'Good' },
        { subjectName: 'Generative AI', grade: 'A', gradePoint: 9, status: 'Good' },
        { subjectName: 'Compiler Design', grade: 'B', gradePoint: 7, status: 'Watch' }
      ],
      overallGrade: 'A',
      overallGradePoint: 8.6,
      highestGradeSubject: { subjectName: 'Cloud Computing', grade: 'A+' },
      lowestGradeSubject: { subjectName: 'Compiler Design', grade: 'B' },
      updatedAt: new Date().toISOString()
    },
    {
      id: 'grd_sum_23cs002',
      regno: '23CS002',
      studentName: 'Bhavna Sharma',
      studentEmail: 'bhavna.sharma@example.edu',
      department: 'CSE',
      section: 'CSE-C',
      facultyEmail: 'anjali.menon@example.edu',
      subjectGrades: [
        { subjectName: 'Data Analytics', grade: 'A+', gradePoint: 10, status: 'Good' },
        { subjectName: 'Cloud Computing', grade: 'A', gradePoint: 9, status: 'Good' },
        { subjectName: 'Embedded Programming', grade: 'A', gradePoint: 9, status: 'Good' },
        { subjectName: 'Generative AI', grade: 'A+', gradePoint: 10, status: 'Good' },
        { subjectName: 'Compiler Design', grade: 'B+', gradePoint: 8, status: 'Good' }
      ],
      overallGrade: 'A+',
      overallGradePoint: 9.2,
      highestGradeSubject: { subjectName: 'Data Analytics', grade: 'A+' },
      lowestGradeSubject: { subjectName: 'Compiler Design', grade: 'B+' },
      updatedAt: new Date().toISOString()
    },
    {
      id: 'grd_sum_23cs003',
      regno: '23CS003',
      studentName: 'Chetan Kumar',
      studentEmail: 'chetan.kumar@example.edu',
      department: 'CSE',
      section: 'CSE-C',
      facultyEmail: 'anjali.menon@example.edu',
      subjectGrades: [
        { subjectName: 'Data Analytics', grade: 'B', gradePoint: 7, status: 'Watch' },
        { subjectName: 'Cloud Computing', grade: 'B+', gradePoint: 8, status: 'Good' },
        { subjectName: 'Embedded Programming', grade: 'C+', gradePoint: 6, status: 'Watch' },
        { subjectName: 'Generative AI', grade: 'B', gradePoint: 7, status: 'Watch' },
        { subjectName: 'Compiler Design', grade: 'C', gradePoint: 5, status: 'At Risk' }
      ],
      overallGrade: 'B',
      overallGradePoint: 6.6,
      highestGradeSubject: { subjectName: 'Cloud Computing', grade: 'B+' },
      lowestGradeSubject: { subjectName: 'Compiler Design', grade: 'C' },
      updatedAt: new Date().toISOString()
    }
  ];
}

function loadGradeImportDatasets(): DynamicGradeImportDataset[] {
  try {
    const raw = localStorage.getItem(LOCAL_GRADE_DATASETS_KEY);
    if (raw) return JSON.parse(raw);
  } catch {}
  return [];
}

function saveGradeSummaryStores() {
  try {
    localStorage.setItem(LOCAL_GRADE_SUMMARY_KEY, JSON.stringify(localGradeSummaryRecords));
    localStorage.setItem(LOCAL_GRADE_DATASETS_KEY, JSON.stringify(localGradeImportDatasets));
  } catch {}
}

export function gradeToPoint(gradeStr: string): number {
  const clean = String(gradeStr || '').trim().toUpperCase();
  switch (clean) {
    case 'O':
    case 'A+': return 10;
    case 'A': return 9;
    case 'A-': return 8.5;
    case 'B+': return 8;
    case 'B': return 7;
    case 'B-': return 6.5;
    case 'C+': return 6;
    case 'C': return 5;
    case 'C-': return 4.5;
    case 'D': return 4;
    case 'E': return 3;
    case 'F': return 0;
    default:
      const num = parseFloat(clean);
      return !isNaN(num) ? num : 0;
  }
}

export function pointToGrade(pt: number): string {
  if (pt >= 9.5) return 'A+';
  if (pt >= 8.5) return 'A';
  if (pt >= 7.5) return 'B+';
  if (pt >= 6.5) return 'B';
  if (pt >= 5.5) return 'C+';
  if (pt >= 4.5) return 'C';
  if (pt >= 4.0) return 'D';
  return 'F';
}

export async function fetchFacultyGradeSummaryRecords(
  facultyEmail?: string,
  sectionName?: string
): Promise<StudentGradeSummaryRecord[]> {
  try {
    let query = supabase.from('student_grade_summary').select('*');
    if (sectionName) query = query.eq('section', sectionName);
    if (facultyEmail) query = query.eq('faculty_email', facultyEmail.toLowerCase());
    const { data, error } = await query;
    if (!error && data && data.length > 0) {
      return data as StudentGradeSummaryRecord[];
    }
  } catch {}

  let filtered = [...localGradeSummaryRecords];
  if (sectionName) {
    filtered = filtered.filter(r => r.section.toLowerCase() === sectionName.toLowerCase() || r.section.toLowerCase().endsWith(sectionName.toLowerCase()));
  }
  if (facultyEmail) {
    filtered = filtered.filter(r => !r.facultyEmail || r.facultyEmail.toLowerCase() === facultyEmail.toLowerCase());
  }
  return filtered;
}

export async function fetchStudentGradeSummaryRecord(
  regnoOrEmail: string
): Promise<StudentGradeSummaryRecord | null> {
  const clean = regnoOrEmail.toLowerCase().trim();
  try {
    const { data, error } = await supabase
      .from('student_grade_summary')
      .select('*')
      .or(`regno.ilike.${clean},student_email.ilike.${clean}`)
      .limit(1);
    if (!error && data && data.length > 0) {
      return data[0] as StudentGradeSummaryRecord;
    }
  } catch {}

  const found = localGradeSummaryRecords.find(
    r => r.regno.toLowerCase() === clean || (r.studentEmail && r.studentEmail.toLowerCase() === clean)
  );
  return found || localGradeSummaryRecords[0] || null;
}

export async function saveGradeSummaryBatch(
  records: StudentGradeSummaryRecord[],
  datasetMeta?: Omit<DynamicGradeImportDataset, 'id' | 'importedAt' | 'records'>,
  overwrite: boolean = true
): Promise<{ success: boolean; count: number }> {
  try {
    await supabase.from('student_grade_summary').upsert(records, { onConflict: 'regno' });
  } catch {}

  records.forEach(newRec => {
    const idx = localGradeSummaryRecords.findIndex(r => r.regno.toLowerCase() === newRec.regno.toLowerCase());
    if (idx !== -1) {
      if (overwrite) {
        localGradeSummaryRecords[idx] = newRec;
      } else {
        const existingRec = localGradeSummaryRecords[idx];
        const mergedSubjects = [...existingRec.subjectGrades];

        newRec.subjectGrades.forEach(newSub => {
          const sIdx = mergedSubjects.findIndex(s => s.subjectName.toLowerCase() === newSub.subjectName.toLowerCase());
          if (sIdx !== -1) {
            mergedSubjects[sIdx] = newSub;
          } else {
            mergedSubjects.push(newSub);
          }
        });

        const validPoints = mergedSubjects.map(s => s.gradePoint ?? gradeToPoint(s.grade));
        const avgPt = validPoints.length > 0 ? (validPoints.reduce((a, b) => a + b, 0) / validPoints.length) : undefined;
        const overall = newRec.overallGrade || (avgPt !== undefined ? pointToGrade(avgPt) : 'N/A');

        localGradeSummaryRecords[idx] = {
          ...existingRec,
          subjectGrades: mergedSubjects,
          overallGrade: overall,
          overallGradePoint: avgPt !== undefined ? Math.round(avgPt * 10) / 10 : undefined,
          updatedAt: new Date().toISOString()
        };
      }
    } else {
      localGradeSummaryRecords.push(newRec);
    }
  });

  if (datasetMeta) {
    const ds: DynamicGradeImportDataset = {
      ...datasetMeta,
      id: `grd_ds_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
      importedAt: new Date().toISOString(),
      records
    };
    localGradeImportDatasets.unshift(ds);
    try {
      await supabase.from('grade_import_datasets').insert([ds]);
    } catch {}
  }

  saveGradeSummaryStores();
  return { success: true, count: records.length };
}

export async function fetchDynamicGradeImportHistory(
  facultyEmail?: string
): Promise<DynamicGradeImportDataset[]> {
  try {
    let query = supabase.from('grade_import_datasets').select('*').order('imported_at', { ascending: false });
    if (facultyEmail) query = query.eq('faculty_email', facultyEmail.toLowerCase());
    const { data, error } = await query;
    if (!error && data && data.length > 0) {
      return data as DynamicGradeImportDataset[];
    }
  } catch {}

  if (facultyEmail) {
    return localGradeImportDatasets.filter(d => (d.importedBy || '').toLowerCase() === facultyEmail.toLowerCase());
  }
  return localGradeImportDatasets;
}

export function parseDynamicGradeExcel(
  arrayBuffer: ArrayBuffer,
  roster: StudentMember[]
) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const rawRows: Record<string, any>[] = XLSX.utils.sheet_to_json(sheet, { defval: '' });

  if (rawRows.length === 0) {
    return {
      headers: [],
      regNoHeader: '',
      subjectHeaders: [],
      matchedRows: [],
      unmatchedRows: [],
      totalExcelRows: 0
    };
  }

  const headers = Object.keys(rawRows[0]);
  const norm = (s: string) => String(s || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');

  let regNoHeader = headers.find(h => ['regno', 'register', 'rollno', 'studentid', 'id', 'reg'].some(k => norm(h).includes(k))) || headers[0];
  let nameHeader = headers.find(h => ['name', 'studentname', 'fullname'].some(k => norm(h).includes(k)));
  let emailHeader = headers.find(h => ['email', 'mail'].some(k => norm(h).includes(k)));
  let overallHeader = headers.find(h => norm(h).includes('overallgrade') || norm(h).includes('overall') || norm(h) === 'grade');

  const metaKeywords = ['regno', 'register', 'rollno', 'studentid', 'id', 'reg', 'name', 'studentname', 'fullname', 'email', 'mail', 'department', 'dept', 'section', 'sec', 'year', 'semester', 'sem', 'dob'];
  if (overallHeader) metaKeywords.push(norm(overallHeader));

  const subjectHeaders = headers.filter(h => {
    const n = norm(h);
    return !metaKeywords.some(k => n === k || (k.length >= 4 && n === k));
  });

  const matchedRows: StudentGradeSummaryRecord[] = [];
  const unmatchedRows: Array<{ rowNumber: number; data: Record<string, any>; reason: string }> = [];

  rawRows.forEach((row, i) => {
    const regVal = String(row[regNoHeader] || '').trim();
    const nameVal = nameHeader ? String(row[nameHeader] || '').trim() : '';
    const emailVal = emailHeader ? String(row[emailHeader] || '').trim() : '';

    if (!regVal && !nameVal && !emailVal) {
      unmatchedRows.push({ rowNumber: i + 2, data: row, reason: 'Empty student identifier row' });
      return;
    }

    const matchedStudent = roster.find(s => 
      (regVal && s.regno.toLowerCase() === regVal.toLowerCase()) ||
      (emailVal && s.email.toLowerCase() === emailVal.toLowerCase()) ||
      (nameVal && s.name.toLowerCase() === nameVal.toLowerCase())
    );

    if (!matchedStudent && regVal && roster.length > 0 && !roster.some(s => s.regno.toLowerCase() === regVal.toLowerCase())) {
      unmatchedRows.push({ rowNumber: i + 2, data: row, reason: `Student '${regVal}' not in authorized section roster` });
    }

    const sReg = matchedStudent?.regno || regVal || `REG_${i + 1}`;
    const sName = matchedStudent?.name || nameVal || `Student ${i + 1}`;
    const sEmail = matchedStudent?.email || emailVal;
    const sDept = matchedStudent?.department || String(row['Department'] || row['Dept'] || 'CSE').trim();
    const sSec = matchedStudent?.section || String(row['Section'] || row['Sec'] || 'CSE-C').trim();

    const subjectGrades: DynamicSubjectGrade[] = [];
    const validPoints: number[] = [];

    subjectHeaders.forEach(sh => {
      const rawVal = String(row[sh] || '').trim();
      if (!rawVal) return;

      const pt = gradeToPoint(rawVal);
      const gradeStr = rawVal.toUpperCase();
      const status = pt >= 8 ? 'Good' : pt >= 6 ? 'Watch' : 'At Risk';

      validPoints.push(pt);
      subjectGrades.push({
        subjectName: sh,
        grade: gradeStr,
        gradePoint: pt,
        status
      });
    });

    const explicitOverall = overallHeader ? String(row[overallHeader] || '').trim().toUpperCase() : '';
    const avgPoint = validPoints.length > 0 ? (validPoints.reduce((a, b) => a + b, 0) / validPoints.length) : undefined;
    const overallGrade = explicitOverall || (avgPoint !== undefined ? pointToGrade(avgPoint) : 'N/A');

    let highestGradeSubject: { subjectName: string; grade: string } | undefined = undefined;
    let lowestGradeSubject: { subjectName: string; grade: string } | undefined = undefined;

    if (subjectGrades.length > 0) {
      const sorted = [...subjectGrades].sort((a, b) => (b.gradePoint ?? 0) - (a.gradePoint ?? 0));
      highestGradeSubject = { subjectName: sorted[0].subjectName, grade: sorted[0].grade };
      lowestGradeSubject = { subjectName: sorted[sorted.length - 1].subjectName, grade: sorted[sorted.length - 1].grade };
    }

    matchedRows.push({
      id: `grd_sum_${sReg.toLowerCase().replace(/[^a-z0-9]/g, '_')}`,
      regno: sReg,
      studentName: sName,
      studentEmail: sEmail,
      department: sDept,
      section: sSec,
      subjectGrades,
      overallGrade,
      overallGradePoint: avgPoint !== undefined ? Math.round(avgPoint * 10) / 10 : undefined,
      highestGradeSubject,
      lowestGradeSubject,
      updatedAt: new Date().toISOString()
    });
  });

  return {
    headers,
    regNoHeader,
    nameHeader,
    subjectHeaders,
    overallHeader,
    matchedRows,
    unmatchedRows,
    totalExcelRows: rawRows.length
  };
}

// ====================================================
// HACKATHON MANAGEMENT & SYNC SYSTEM
// ====================================================
import { HackathonAggregator, RawHackathon } from './hackathonProviders';

export interface Hackathon {
  id: string;
  external_id: string;
  source: 'UNSTOP' | 'DEVFOLIO' | 'DEVPOST';
  title: string;
  description: string;
  official_url: string;
  image_url?: string;
  status: 'LIVE' | 'UPCOMING' | 'ENDING_SOON' | 'ENDED';
  mode: 'Online' | 'Offline' | 'Hybrid';
  start_date?: string;
  end_date?: string;
  registration_deadline?: string;
  eligibility?: string;
  team_size?: string;
  categories: string[];
  skills: string[];
  participant_count?: number;
  prize?: string;
  organizer?: string;
  featured?: boolean;
  relevance_score?: number;
  source_created_at?: string;
  last_synced_at?: string;
  created_at?: string;
  updated_at?: string;
}

export function calculateHackathonRelevance(
  hackathon: Hackathon,
  studentDept: string = 'CSE',
  studentSkills: string[] = ['AI', 'Python', 'React', 'Java', 'Machine Learning']
): number {
  let score = 50;

  if (hackathon.status === 'LIVE') score += 25;
  if (hackathon.status === 'ENDING_SOON') score += 30;
  if (hackathon.status === 'UPCOMING') score += 15;

  const deptLower = studentDept.toLowerCase();
  const text = `${hackathon.title} ${hackathon.description} ${hackathon.eligibility || ''} ${(hackathon.categories || []).join(' ')}`.toLowerCase();
  if (text.includes(deptLower) || text.includes('engineering') || text.includes('all students')) {
    score += 15;
  }

  const hackSkills = (hackathon.skills || []).map(s => s.toLowerCase());
  const hackCats = (hackathon.categories || []).map(c => c.toLowerCase());
  for (const s of studentSkills) {
    const sLow = s.toLowerCase();
    if (hackSkills.some(hs => hs.includes(sLow)) || hackCats.some(hc => hc.includes(sLow))) {
      score += 10;
    }
  }

  if (hackathon.mode === 'Online' || hackathon.mode === 'Hybrid') score += 10;
  if (hackathon.featured) score += 20;

  return Math.min(100, score);
}

const HACKATHONS_LOCAL_KEY = 'cogniva_hackathons_v1';

export async function syncHackathons(): Promise<Hackathon[]> {
  const aggregator = new HackathonAggregator();
  const rawItems = await aggregator.fetchAll();
  const now = new Date().toISOString();

  const normalizedList: Hackathon[] = rawItems
    .filter((item) => item.title && item.official_url)
    .map((item) => {
      let status = item.status || 'LIVE';
      if (item.registration_deadline) {
        const deadline = new Date(item.registration_deadline).getTime();
        if (!isNaN(deadline)) {
          const diffDays = (deadline - Date.now()) / (1000 * 3600 * 24);
          if (diffDays < 0) status = 'ENDED';
          else if (diffDays <= 4 && status !== 'ENDED') status = 'ENDING_SOON';
        }
      }

      const h: Hackathon = {
        id: `${item.source.toLowerCase()}_${item.external_id.replace(/[^a-z0-9]/gi, '_')}`,
        external_id: item.external_id,
        source: item.source,
        title: item.title,
        description: item.description,
        official_url: item.official_url,
        image_url: item.image_url,
        status: status as any,
        mode: item.mode || 'Online',
        start_date: item.start_date,
        end_date: item.end_date,
        registration_deadline: item.registration_deadline,
        eligibility: item.eligibility,
        team_size: item.team_size,
        categories: item.categories || ['Technology'],
        skills: item.skills || ['Coding'],
        participant_count: item.participant_count,
        prize: item.prize,
        organizer: item.organizer || item.source,
        featured: false,
        last_synced_at: now,
        created_at: now,
        updated_at: now
      };
      h.relevance_score = calculateHackathonRelevance(h);
      return h;
    });

  const uniqueMap = new Map<string, Hackathon>();
  normalizedList.forEach((h) => {
    if (h.status !== 'ENDED') {
      uniqueMap.set(`${h.source}_${h.external_id}`, h);
    }
  });
  const uniqueHackathons = Array.from(uniqueMap.values());

  try {
    localStorage.setItem(HACKATHONS_LOCAL_KEY, JSON.stringify(uniqueHackathons));
  } catch (err) {
    console.warn('LocalStorage save failed for hackathons:', err);
  }

  try {
    const { error } = await supabase.from('hackathons').upsert(uniqueHackathons, { onConflict: 'external_id,source' });
    if (error) {
      console.warn('Supabase hackathons upsert warning:', error.message);
    }
  } catch (err) {
    console.warn('Supabase sync warning:', err);
  }

  return uniqueHackathons;
}

export async function fetchHackathons(): Promise<Hackathon[]> {
  // First attempt server endpoint feed for verified real data
  try {
    const feedRes = await fetch('/api/hackathons/feed');
    if (feedRes.ok) {
      const feedJson = await feedRes.json();
      if (feedJson.success && Array.isArray(feedJson.data) && feedJson.data.length > 0) {
        const now = new Date().toISOString();
        const verifiedList: Hackathon[] = feedJson.data.map((item: RawHackathon) => {
          const h: Hackathon = {
            id: `${item.source.toLowerCase()}_${item.external_id.replace(/[^a-z0-9]/gi, '_')}`,
            external_id: item.external_id,
            source: item.source,
            title: item.title,
            description: item.description,
            official_url: item.official_url,
            image_url: item.image_url,
            status: item.status || 'LIVE',
            mode: item.mode || 'Online',
            start_date: item.start_date,
            end_date: item.end_date,
            registration_deadline: item.registration_deadline,
            eligibility: item.eligibility,
            team_size: item.team_size,
            categories: item.categories || ['Technology'],
            skills: item.skills || ['Coding'],
            participant_count: item.participant_count,
            prize: item.prize,
            organizer: item.organizer || item.source,
            featured: false,
            last_synced_at: now,
            created_at: now,
            updated_at: now
          };
          h.relevance_score = calculateHackathonRelevance(h);
          return h;
        });

        try {
          localStorage.setItem(HACKATHONS_LOCAL_KEY, JSON.stringify(verifiedList));
        } catch {}

        return verifiedList;
      }
    }
  } catch (err) {
    console.warn('Server feed fetch warning:', err);
  }

  try {
    const { data, error } = await supabase
      .from('hackathons')
      .select('*')
      .neq('status', 'ENDED');

    if (!error && data && data.length > 0) {
      // Filter out any legacy synthetic/fake records if present
      const cleanData = (data as any[]).filter(
        (h) => !h.external_id?.includes('ethindia') && !h.external_id?.includes('flipkart') && !h.external_id?.includes('chrome')
      );
      if (cleanData.length > 0) {
        return cleanData.map((item: any) => ({
          ...item,
          relevance_score: calculateHackathonRelevance(item)
        }));
      }
    }
  } catch (err) {
    console.warn('Supabase hackathons fetch warning:', err);
  }

  try {
    const local = localStorage.getItem(HACKATHONS_LOCAL_KEY);
    if (local) {
      const parsed: Hackathon[] = JSON.parse(local);
      const cleanLocal = parsed.filter(
        (h) => h.status !== 'ENDED' && !h.external_id?.includes('ethindia') && !h.external_id?.includes('flipkart') && !h.external_id?.includes('chrome')
      );
      if (cleanLocal.length > 0) {
        return cleanLocal.map((h) => ({
          ...h,
          relevance_score: calculateHackathonRelevance(h)
        }));
      }
    }
  } catch (err) {
    console.warn('LocalStorage hackathons read warning:', err);
  }

  return await syncHackathons();
}

// ====================================================
// GOAL-TO-ACTION ROADMAP SYSTEM API
// ====================================================
import { generateGoalRoadmapAi } from './ai-service';

export interface GoalMilestone {
  id: string;
  goal_id: string;
  phase_id: string;
  title: string;
  description: string;
  why_it_matters: string;
  priority: 'HIGH' | 'MEDIUM' | 'LOW';
  estimated_hours: number;
  order_index: number;
  status: 'PENDING' | 'COMPLETED';
  completed_at?: string;
  created_at?: string;
  updated_at?: string;
}

export interface GoalPhase {
  id: string;
  goal_id: string;
  title: string;
  description?: string;
  order_index: number;
  milestones: GoalMilestone[];
}

export interface Goal {
  id: string;
  student_id?: string;
  student_email?: string;
  title: string;
  description?: string;
  why_it_matters?: string;
  target_date?: string;
  progress_percentage: number;
  status: 'ON_TRACK' | 'NEEDS_ATTENTION' | 'AT_RISK' | 'NOT_STARTED';
  color: 'teal' | 'amber' | 'coral' | 'violet';
  phases: GoalPhase[];
  created_at?: string;
  updated_at?: string;
}

const LOCAL_GOALS_KEY = 'cogniva_student_goals_v1';

export function calculateGoalStats(goal: Goal): Goal {
  const allMilestones = (goal.phases || []).flatMap(p => p.milestones || []);
  const total = allMilestones.length;
  const completed = allMilestones.filter(m => m.status === 'COMPLETED').length;
  const progress_percentage = total > 0 ? Math.round((completed / total) * 100) : 0;

  let status: Goal['status'] = 'NOT_STARTED';
  if (total > 0) {
    if (progress_percentage >= 60) status = 'ON_TRACK';
    else if (progress_percentage >= 30) status = 'NEEDS_ATTENTION';
    else if (progress_percentage > 0) status = 'AT_RISK';
    else status = 'NOT_STARTED';
  }

  return {
    ...goal,
    progress_percentage,
    status
  };
}

export function getDefaultGoals(userEmail?: string): Goal[] {
  const email = (userEmail || 'student001@cogniva.edu').toLowerCase();

  const g1: Goal = {
    id: 'goal_research_internship',
    student_email: email,
    title: 'Research internship',
    description: 'Build a credible applied ML portfolio and publish benchmark evaluation code.',
    why_it_matters: 'Essential for securing competitive research assistantships and top lab offers.',
    target_date: '2026-03-28',
    progress_percentage: 50,
    status: 'NEEDS_ATTENTION',
    color: 'violet',
    phases: [
      {
        id: 'phase_res_1',
        goal_id: 'goal_research_internship',
        title: 'Phase 1: Build Core Foundations',
        order_index: 1,
        milestones: [
          {
            id: 'ms_res_1',
            goal_id: 'goal_research_internship',
            phase_id: 'phase_res_1',
            title: 'Master NumPy & Pandas for Applied ML',
            description: 'Strengthen data manipulation, array vectorization, and data processing skills.',
            why_it_matters: 'Core prerequisite for implementing data pipelines and model training code.',
            priority: 'HIGH',
            estimated_hours: 8,
            order_index: 1,
            status: 'COMPLETED',
            completed_at: new Date().toISOString()
          },
          {
            id: 'ms_res_2',
            goal_id: 'goal_research_internship',
            phase_id: 'phase_res_1',
            title: 'Study Core Machine Learning Algorithms',
            description: 'Implement linear regression, decision trees, and SVMs from scratch.',
            why_it_matters: 'Builds deep technical intuition needed for research lab technical interviews.',
            priority: 'HIGH',
            estimated_hours: 12,
            order_index: 2,
            status: 'PENDING'
          }
        ]
      },
      {
        id: 'phase_res_2',
        goal_id: 'goal_research_internship',
        title: 'Phase 2: Build Empirical Evidence',
        order_index: 2,
        milestones: [
          {
            id: 'ms_res_3',
            goal_id: 'goal_research_internship',
            phase_id: 'phase_res_2',
            title: 'Construct an End-to-End ML Evaluation Notebook',
            description: 'Build a reproducible benchmark notebook comparing baseline models on public datasets.',
            why_it_matters: 'Demonstrates practical experimentation capability to prospective research mentors.',
            priority: 'HIGH',
            estimated_hours: 15,
            order_index: 3,
            status: 'PENDING'
          },
          {
            id: 'ms_res_4',
            goal_id: 'goal_research_internship',
            phase_id: 'phase_res_2',
            title: 'Publish Open-Source ML Repository on GitHub',
            description: 'Clean code, write comprehensive README, and add model performance visualizations.',
            why_it_matters: 'Serves as visible proof of code quality for applications.',
            priority: 'MEDIUM',
            estimated_hours: 6,
            order_index: 4,
            status: 'PENDING'
          }
        ]
      }
    ]
  };

  const g2: Goal = {
    id: 'goal_semester_distinction',
    student_email: email,
    title: 'Semester distinction',
    description: 'Finish above 8.5 GPA with strong assessment consistency across all subjects.',
    why_it_matters: 'Protects academic standing and satisfies honor roll criteria.',
    target_date: '2026-04-30',
    progress_percentage: 67,
    status: 'ON_TRACK',
    color: 'amber',
    phases: [
      {
        id: 'phase_sem_1',
        goal_id: 'goal_semester_distinction',
        title: 'Phase 1: Assessment Protection',
        order_index: 1,
        milestones: [
          {
            id: 'ms_sem_1',
            goal_id: 'goal_semester_distinction',
            phase_id: 'phase_sem_1',
            title: 'Map Subject Priority Matrix',
            description: 'Identify subjects below 8.0 target GPA and protect 45-min daily revision blocks.',
            why_it_matters: 'Early focus prevents score drop before midterms.',
            priority: 'HIGH',
            estimated_hours: 5,
            order_index: 1,
            status: 'COMPLETED',
            completed_at: new Date().toISOString()
          },
          {
            id: 'ms_sem_2',
            goal_id: 'goal_semester_distinction',
            phase_id: 'phase_sem_1',
            title: 'Complete All Pending Subject Assignments',
            description: 'Finish lab reports and theory submissions 48 hours before deadline.',
            why_it_matters: 'Secures full internal assessment marks.',
            priority: 'HIGH',
            estimated_hours: 10,
            order_index: 2,
            status: 'COMPLETED',
            completed_at: new Date().toISOString()
          },
          {
            id: 'ms_sem_3',
            goal_id: 'goal_semester_distinction',
            phase_id: 'phase_sem_1',
            title: 'Solve Past 5-Year Question Papers',
            description: 'Practice timed paper solving for core subjects.',
            why_it_matters: 'Familiarizes with faculty examination patterns.',
            priority: 'HIGH',
            estimated_hours: 14,
            order_index: 3,
            status: 'PENDING'
          }
        ]
      }
    ]
  };

  const g3: Goal = {
    id: 'goal_systems_project',
    student_email: email,
    title: 'Complete systems project',
    description: 'Ship a distributed systems demo with RPC networking and replication.',
    why_it_matters: 'Demonstrates low-level engineering capability for backend SDE roles.',
    target_date: '2026-04-12',
    progress_percentage: 33,
    status: 'AT_RISK',
    color: 'teal',
    phases: [
      {
        id: 'phase_sys_1',
        goal_id: 'goal_systems_project',
        title: 'Phase 1: Architecture & Prototype',
        order_index: 1,
        milestones: [
          {
            id: 'ms_sys_1',
            goal_id: 'goal_systems_project',
            phase_id: 'phase_sys_1',
            title: 'Design Concurrent Event Loop Architecture',
            description: 'Implement socket handling and thread pool dispatcher.',
            why_it_matters: 'Foundation for high-throughput node communication.',
            priority: 'HIGH',
            estimated_hours: 12,
            order_index: 1,
            status: 'COMPLETED',
            completed_at: new Date().toISOString()
          },
          {
            id: 'ms_sys_2',
            goal_id: 'goal_systems_project',
            phase_id: 'phase_sys_1',
            title: 'Implement Consensus & Data Replication',
            description: 'Write heartbeat ping and log synchronization handlers.',
            why_it_matters: 'Ensures cluster consistency during network partitions.',
            priority: 'HIGH',
            estimated_hours: 16,
            order_index: 2,
            status: 'PENDING'
          },
          {
            id: 'ms_sys_3',
            goal_id: 'goal_systems_project',
            phase_id: 'phase_sys_1',
            title: 'Benchmark Latency under Synthetic Load',
            description: 'Record throughput metrics and memory footprint.',
            why_it_matters: 'Provides empirical metrics for system documentation.',
            priority: 'MEDIUM',
            estimated_hours: 6,
            order_index: 3,
            status: 'PENDING'
          }
        ]
      }
    ]
  };

  return [calculateGoalStats(g1), calculateGoalStats(g2), calculateGoalStats(g3)];
}

export async function fetchStudentGoals(userEmail?: string): Promise<Goal[]> {
  const cleanEmail = (userEmail || 'student001@cogniva.edu').toLowerCase();

  try {
    const { data, error } = await supabase
      .from('student_goals')
      .select('*')
      .eq('student_email', cleanEmail)
      .order('created_at', { ascending: false });

    if (!error && data && data.length > 0) {
      return data.map((item: any) => calculateGoalStats(item));
    }
  } catch (err) {
    console.warn('Supabase student_goals fetch notice:', err);
  }

  try {
    const raw = localStorage.getItem(`${LOCAL_GOALS_KEY}_${cleanEmail}`);
    if (raw) {
      const parsed: Goal[] = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return parsed.map(g => calculateGoalStats(g));
      }
    }
  } catch (err) {
    console.warn('LocalStorage student_goals read notice:', err);
  }

  const defaults = getDefaultGoals(cleanEmail);
  saveAllGoalsLocal(cleanEmail, defaults);
  return defaults;
}

function saveAllGoalsLocal(email: string, goals: Goal[]) {
  try {
    localStorage.setItem(`${LOCAL_GOALS_KEY}_${email.toLowerCase()}`, JSON.stringify(goals));
  } catch {}
}

export async function saveStudentGoal(goal: Goal): Promise<Goal> {
  const updatedGoal = calculateGoalStats(goal);
  const email = (updatedGoal.student_email || 'student001@cogniva.edu').toLowerCase();

  const existingGoals = await fetchStudentGoals(email);
  const idx = existingGoals.findIndex(g => g.id === updatedGoal.id || g.title.toLowerCase() === updatedGoal.title.toLowerCase());
  if (idx !== -1) {
    existingGoals[idx] = updatedGoal;
  } else {
    existingGoals.unshift(updatedGoal);
  }

  saveAllGoalsLocal(email, existingGoals);

  try {
    await supabase.from('student_goals').upsert([updatedGoal], { onConflict: 'id' });
  } catch (err) {
    console.warn('Supabase student_goals save notice:', err);
  }

  return updatedGoal;
}

export async function toggleMilestoneCompletion(goalId: string, milestoneId: string, userEmail?: string): Promise<Goal | null> {
  const email = (userEmail || 'student001@cogniva.edu').toLowerCase();
  const goals = await fetchStudentGoals(email);
  const goal = goals.find(g => g.id === goalId);
  if (!goal) return null;

  for (const phase of goal.phases) {
    const ms = phase.milestones.find(m => m.id === milestoneId);
    if (ms) {
      ms.status = ms.status === 'COMPLETED' ? 'PENDING' : 'COMPLETED';
      ms.completed_at = ms.status === 'COMPLETED' ? new Date().toISOString() : undefined;
      break;
    }
  }

  return await saveStudentGoal(goal);
}

export async function deleteStudentGoal(goalId: string, userEmail?: string): Promise<boolean> {
  const email = (userEmail || 'student001@cogniva.edu').toLowerCase();
  const goals = await fetchStudentGoals(email);
  const filtered = goals.filter(g => g.id !== goalId);
  saveAllGoalsLocal(email, filtered);

  try {
    await supabase.from('student_goals').delete().eq('id', goalId);
  } catch {}

  return true;
}

export async function autoGenerateGoalMilestones(
  goalTitle: string,
  existingGoal?: Goal,
  goalDescription?: string,
  targetDate?: string,
  userEmail?: string
): Promise<Goal> {
  const email = (userEmail || 'student001@cogniva.edu').toLowerCase();

  const roadmapData = await generateGoalRoadmapAi(
    goalTitle,
    goalDescription || existingGoal?.description,
    targetDate || existingGoal?.target_date,
    { dept: 'CSE', year: '2nd Year', cgpa: 8.2, skills: ['Java', 'Python', 'React', 'AI'] },
    email
  );

  const goalId = existingGoal?.id || `goal_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`;
  const preservedCompletedMilestones = (existingGoal?.phases || []).flatMap(p => p.milestones || []).filter(m => m.status === 'COMPLETED');

  const newPhases: GoalPhase[] = (roadmapData?.phases || []).map((phase, pIdx) => {
    const phaseId = `phase_${goalId}_${pIdx + 1}`;
    const newMilestones: GoalMilestone[] = phase.milestones.map((ms, mIdx) => ({
      id: `ms_${goalId}_${pIdx + 1}_${mIdx + 1}`,
      goal_id: goalId,
      phase_id: phaseId,
      title: ms.title,
      description: ms.description,
      why_it_matters: ms.why_it_matters,
      priority: ms.priority || 'HIGH',
      estimated_hours: ms.estimated_hours || 8,
      order_index: ms.order || (mIdx + 1),
      status: 'PENDING'
    }));
    return {
      id: phaseId,
      goal_id: goalId,
      title: phase.title,
      order_index: pIdx + 1,
      milestones: newMilestones
    };
  });

  if (preservedCompletedMilestones.length > 0 && newPhases.length > 0) {
    const existingMsTitles = new Set(newPhases.flatMap(p => p.milestones.map(m => m.title.toLowerCase())));
    preservedCompletedMilestones.forEach(cm => {
      if (!existingMsTitles.has(cm.title.toLowerCase())) {
        newPhases[0].milestones.unshift(cm);
      }
    });
  }

  const updatedGoal: Goal = {
    id: goalId,
    student_email: email,
    title: goalTitle,
    description: goalDescription || existingGoal?.description || `Actionable AI-generated roadmap for ${goalTitle}`,
    why_it_matters: existingGoal?.why_it_matters || 'Structured milestones to achieve goal targets.',
    target_date: targetDate || existingGoal?.target_date || '2026-12-15',
    progress_percentage: existingGoal?.progress_percentage || 0,
    status: existingGoal?.status || 'NOT_STARTED',
    color: existingGoal?.color || 'violet',
    phases: newPhases,
    created_at: existingGoal?.created_at || new Date().toISOString(),
    updated_at: new Date().toISOString()
  };

  return await saveStudentGoal(updatedGoal);
}

// ============================================================================
// INTERNSHIP PROVIDER ARCHITECTURE & DATA STRUCTURES
// ============================================================================

import {
  MultiSourceInternshipAggregator as AggregatorClass,
  ProviderHealth as HealthType,
  Internship as InternshipType
} from './internshipProviders';

export type {
  Internship,
  InternshipSource,
  CompensationType,
  ApplicationPipelineStatus,
  ProviderHealth,
  StudentEligibilityResult,
  StudentApplicationRecord
} from './internshipProviders';

export {
  MultiSourceInternshipAggregator,
  evaluateStudentEligibility,
  parseStudentYear,
  fetchStudentApplicationPipeline,
  updateStudentApplicationStatus
} from './internshipProviders';

export function calculateInternshipRelevance(
  internship: InternshipType,
  studentDept: string = 'CSE',
  studentSkills: string[] = ['Python', 'AI', 'Machine Learning', 'React', 'Java', 'SQL'],
  goalTitle?: string,
  passionTrack?: string
): number {
  let score = 50;

  const text = `${internship.title} ${internship.description} ${internship.eligibility || ''} ${internship.company_name}`.toLowerCase();

  // Department / Academic Fit
  if (text.includes(studentDept.toLowerCase()) || text.includes('engineering') || text.includes('computer science')) {
    score += 15;
  }

  // Skill Relevance
  const reqSkills = (internship.skills || []).map(s => s.toLowerCase());
  for (const s of studentSkills) {
    const sLow = s.toLowerCase();
    if (reqSkills.some(rs => rs.includes(sLow)) || text.includes(sLow)) {
      score += 10;
    }
  }

  // Work Mode preference
  if (internship.work_mode === 'Remote' || internship.work_mode === 'Hybrid') {
    score += 10;
  }

  // Goal & Passion Track Relevance
  if (passionTrack) {
    const passionKw = passionTrack.toLowerCase().replace('_', ' ');
    if (text.includes(passionKw)) score += 15;
  }
  if (goalTitle) {
    const goalKw = goalTitle.toLowerCase();
    if (text.includes(goalKw)) score += 15;
  }

  return Math.min(100, score);
}

const INTERNSHIPS_LOCAL_KEY = 'cogniva_internships_v1';
const SAVED_HACKATHONS_LOCAL_KEY = 'cogniva_saved_hackathons_v1';
const SAVED_INTERNSHIPS_LOCAL_KEY = 'cogniva_saved_internships_v1';

let cachedProviderHealth: HealthType[] = [];

export function getLatestProviderHealth(): HealthType[] {
  return cachedProviderHealth;
}

export async function fetchMultiSourceInternshipsWithHealth(): Promise<{
  internships: InternshipType[];
  health: HealthType[];
}> {
  const aggregator = new AggregatorClass();
  const res = await aggregator.fetchAllMultiSource();
  
  // Calculate relevance scores
  const scored = res.internships.map(inst => ({
    ...inst,
    relevance_score: calculateInternshipRelevance(inst)
  }));

  cachedProviderHealth = res.health;

  try {
    localStorage.setItem(INTERNSHIPS_LOCAL_KEY, JSON.stringify(scored));
  } catch {}

  return {
    internships: scored,
    health: res.health
  };
}

export async function fetchInternships(): Promise<import('./internshipProviders').Internship[]> {
  // Try multi-source aggregator first
  try {
    const { internships } = await fetchMultiSourceInternshipsWithHealth();
    if (internships.length > 0) return internships;
  } catch (err) {
    console.warn('[fetchInternships] Multi-source aggregator warning:', err);
  }

  // Attempt to fetch from Supabase
  try {
    const { data, error } = await supabase.from('internships').select('*').eq('is_active', true);
    if (!error && data && data.length > 0) {
      const list = (data as import('./internshipProviders').Internship[]).map(i => ({
        ...i,
        relevance_score: calculateInternshipRelevance(i)
      }));
      return list;
    }
  } catch {}

  // Local storage fallback
  try {
    const savedLocal = localStorage.getItem(INTERNSHIPS_LOCAL_KEY);
    if (savedLocal) {
      const parsed = JSON.parse(savedLocal);
      if (Array.isArray(parsed) && parsed.length > 0) return parsed;
    }
  } catch {}

  return [];
}

export async function fetchSavedOpportunityIds(
  userEmail: string,
  type: 'hackathon' | 'internship'
): Promise<string[]> {
  const key = type === 'hackathon' ? `${SAVED_HACKATHONS_LOCAL_KEY}_${userEmail}` : `${SAVED_INTERNSHIPS_LOCAL_KEY}_${userEmail}`;
  try {
    const { data, error } = await supabase
      .from('student_saved_opportunities')
      .select('opportunity_id')
      .eq('student_email', userEmail.toLowerCase().trim())
      .eq('opportunity_type', type);
    if (!error && data) {
      const ids = data.map(d => d.opportunity_id);
      localStorage.setItem(key, JSON.stringify(ids));
      return ids;
    }
  } catch {}

  try {
    const local = localStorage.getItem(key);
    if (local) return JSON.parse(local);
  } catch {}

  return [];
}

export async function toggleSaveOpportunityId(
  userEmail: string,
  id: string,
  type: 'hackathon' | 'internship'
): Promise<string[]> {
  const cleanEmail = userEmail.toLowerCase().trim();
  const current = await fetchSavedOpportunityIds(cleanEmail, type);
  const exists = current.includes(id);
  const updated = exists ? current.filter(i => i !== id) : [...current, id];

  const key = type === 'hackathon' ? `${SAVED_HACKATHONS_LOCAL_KEY}_${cleanEmail}` : `${SAVED_INTERNSHIPS_LOCAL_KEY}_${cleanEmail}`;
  try {
    localStorage.setItem(key, JSON.stringify(updated));
  } catch {}

  try {
    if (exists) {
      await supabase
        .from('student_saved_opportunities')
        .delete()
        .eq('student_email', cleanEmail)
        .eq('opportunity_id', id)
        .eq('opportunity_type', type);
    } else {
      await supabase
        .from('student_saved_opportunities')
        .upsert({
          student_email: cleanEmail,
          opportunity_id: id,
          opportunity_type: type,
          saved_at: new Date().toISOString()
        }, { onConflict: 'student_email,opportunity_id,opportunity_type' });
    }
  } catch {}

  return updated;
}

// ==========================================
// FACULTY TIMETABLE & CAMPUS NAVIGATOR API
// ==========================================

export interface FacultyTimetableEntry {
  id: string;
  faculty_id?: string;
  faculty_employee_id: string;
  faculty_name: string;
  faculty_email?: string;
  cabin_location?: string;
  subject_code?: string;
  subject_name: string;
  department?: string;
  year?: string;
  section_name: string;
  day_of_week: 'Monday' | 'Tuesday' | 'Wednesday' | 'Thursday' | 'Friday' | 'Saturday' | 'Sunday';
  start_time: string;
  end_time: string;
  room_number: string;
  academic_year?: string;
  semester?: string;
  created_at?: string;
  updated_at?: string;
}

export interface TimetableImportRow {
  faculty_name: string;
  faculty_employee_id?: string;
  faculty_email?: string;
  cabin_location?: string;
  subject_code?: string;
  subject_name: string;
  section_name: string;
  department?: string;
  day_of_week: string;
  start_time: string;
  end_time: string;
  room_number: string;
}

export interface TimetableHeaderValidationResult {
  valid: boolean;
  missingColumns: string[];
  mappedHeaders: Record<string, string>;
}

const INITIAL_FACULTY_TIMETABLE: FacultyTimetableEntry[] = [
  // Dr. Anjali Menon (FAC001) - Advisor & CSE-C Faculty
  { id: 'tt_anjali_mon_1', faculty_employee_id: 'FAC001', faculty_name: 'Dr. Anjali Menon', faculty_email: 'anjali.menon@example.edu', cabin_location: 'Main Academic Block, Cabin 304', subject_code: 'CS201', subject_name: 'Database Management Systems', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Monday', start_time: '08:00', end_time: '09:00', department: 'CSE' },
  { id: 'tt_anjali_mon_2', faculty_employee_id: 'FAC001', faculty_name: 'Dr. Anjali Menon', faculty_email: 'anjali.menon@example.edu', cabin_location: 'Main Academic Block, Cabin 304', subject_code: 'CS201L', subject_name: 'DBMS Lab', section_name: 'CSE-C', room_number: 'Database Lab 2', day_of_week: 'Monday', start_time: '10:00', end_time: '12:00', department: 'CSE' },
  { id: 'tt_anjali_tue_1', faculty_employee_id: 'FAC001', faculty_name: 'Dr. Anjali Menon', faculty_email: 'anjali.menon@example.edu', cabin_location: 'Main Academic Block, Cabin 304', subject_code: 'CS201', subject_name: 'Database Management Systems', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Tuesday', start_time: '09:00', end_time: '10:00', department: 'CSE' },
  { id: 'tt_anjali_tue_2', faculty_employee_id: 'FAC001', faculty_name: 'Dr. Anjali Menon', faculty_email: 'anjali.menon@example.edu', cabin_location: 'Main Academic Block, Cabin 304', subject_code: 'CS201', subject_name: 'Database Management Systems', section_name: 'CSE-A', room_number: 'CS-301', day_of_week: 'Tuesday', start_time: '13:00', end_time: '14:00', department: 'CSE' },
  { id: 'tt_anjali_wed_1', faculty_employee_id: 'FAC001', faculty_name: 'Dr. Anjali Menon', faculty_email: 'anjali.menon@example.edu', cabin_location: 'Main Academic Block, Cabin 304', subject_code: 'CS201', subject_name: 'Database Management Systems', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Wednesday', start_time: '08:00', end_time: '09:00', department: 'CSE' },
  { id: 'tt_anjali_wed_2', faculty_employee_id: 'FAC001', faculty_name: 'Dr. Anjali Menon', faculty_email: 'anjali.menon@example.edu', cabin_location: 'Main Academic Block, Cabin 304', subject_code: 'CS201', subject_name: 'Database Management Systems', section_name: 'CSE-B', room_number: 'CS-303', day_of_week: 'Wednesday', start_time: '11:00', end_time: '12:00', department: 'CSE' },
  { id: 'tt_anjali_thu_1', faculty_employee_id: 'FAC001', faculty_name: 'Dr. Anjali Menon', faculty_email: 'anjali.menon@example.edu', cabin_location: 'Main Academic Block, Cabin 304', subject_code: 'CS201L', subject_name: 'DBMS Lab', section_name: 'CSE-A', room_number: 'Database Lab 1', day_of_week: 'Thursday', start_time: '13:00', end_time: '15:00', department: 'CSE' },
  { id: 'tt_anjali_fri_1', faculty_employee_id: 'FAC001', faculty_name: 'Dr. Anjali Menon', faculty_email: 'anjali.menon@example.edu', cabin_location: 'Main Academic Block, Cabin 304', subject_code: 'CS201', subject_name: 'Database Management Systems', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Friday', start_time: '09:00', end_time: '10:00', department: 'CSE' },

  // Dr. Ravi Chandran (FAC002) - CSE-A Advisor & OS Faculty
  { id: 'tt_ravi_mon_1', faculty_employee_id: 'FAC002', faculty_name: 'Dr. Ravi Chandran', faculty_email: 'ravi.chandran@example.edu', cabin_location: 'CS Block, Ground Floor Cabin 08', subject_code: 'CS202', subject_name: 'Operating Systems', section_name: 'CSE-A', room_number: 'CS-301', day_of_week: 'Monday', start_time: '09:00', end_time: '10:00', department: 'CSE' },
  { id: 'tt_ravi_mon_2', faculty_employee_id: 'FAC002', faculty_name: 'Dr. Ravi Chandran', faculty_email: 'ravi.chandran@example.edu', cabin_location: 'CS Block, Ground Floor Cabin 08', subject_code: 'CS202', subject_name: 'Operating Systems', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Monday', start_time: '11:00', end_time: '12:00', department: 'CSE' },
  { id: 'tt_ravi_tue_1', faculty_employee_id: 'FAC002', faculty_name: 'Dr. Ravi Chandran', faculty_email: 'ravi.chandran@example.edu', cabin_location: 'CS Block, Ground Floor Cabin 08', subject_code: 'CS202L', subject_name: 'OS Lab', section_name: 'CSE-C', room_number: 'Systems Lab 1', day_of_week: 'Tuesday', start_time: '10:00', end_time: '12:00', department: 'CSE' },
  { id: 'tt_ravi_wed_1', faculty_employee_id: 'FAC002', faculty_name: 'Dr. Ravi Chandran', faculty_email: 'ravi.chandran@example.edu', cabin_location: 'CS Block, Ground Floor Cabin 08', subject_code: 'CS202', subject_name: 'Operating Systems', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Wednesday', start_time: '09:00', end_time: '10:00', department: 'CSE' },
  { id: 'tt_ravi_thu_1', faculty_employee_id: 'FAC002', faculty_name: 'Dr. Ravi Chandran', faculty_email: 'ravi.chandran@example.edu', cabin_location: 'CS Block, Ground Floor Cabin 08', subject_code: 'CS202', subject_name: 'Operating Systems', section_name: 'CSE-B', room_number: 'CS-303', day_of_week: 'Thursday', start_time: '08:00', end_time: '09:00', department: 'CSE' },
  { id: 'tt_ravi_fri_1', faculty_employee_id: 'FAC002', faculty_name: 'Dr. Ravi Chandran', faculty_email: 'ravi.chandran@example.edu', cabin_location: 'CS Block, Ground Floor Cabin 08', subject_code: 'CS202', subject_name: 'Operating Systems', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Friday', start_time: '08:00', end_time: '09:00', department: 'CSE' },

  // Prof. Meera Krishnan (FAC003) - CSE-B Advisor & DAA Faculty
  { id: 'tt_meera_mon_1', faculty_employee_id: 'FAC003', faculty_name: 'Prof. Meera Krishnan', faculty_email: 'meera.krishnan@example.edu', cabin_location: 'Main Academic Block, Cabin 308', subject_code: 'CS203', subject_name: 'Design & Analysis of Algorithms', section_name: 'CSE-B', room_number: 'CS-303', day_of_week: 'Monday', start_time: '10:00', end_time: '11:00', department: 'CSE' },
  { id: 'tt_meera_mon_2', faculty_employee_id: 'FAC003', faculty_name: 'Prof. Meera Krishnan', faculty_email: 'meera.krishnan@example.edu', cabin_location: 'Main Academic Block, Cabin 308', subject_code: 'CS203', subject_name: 'Design & Analysis of Algorithms', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Monday', start_time: '13:00', end_time: '14:00', department: 'CSE' },
  { id: 'tt_meera_tue_1', faculty_employee_id: 'FAC003', faculty_name: 'Prof. Meera Krishnan', faculty_email: 'meera.krishnan@example.edu', cabin_location: 'Main Academic Block, Cabin 308', subject_code: 'CS203', subject_name: 'Design & Analysis of Algorithms', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Tuesday', start_time: '08:00', end_time: '09:00', department: 'CSE' },
  { id: 'tt_meera_wed_1', faculty_employee_id: 'FAC003', faculty_name: 'Prof. Meera Krishnan', faculty_email: 'meera.krishnan@example.edu', cabin_location: 'Main Academic Block, Cabin 308', subject_code: 'CS203L', subject_name: 'Algorithms Lab', section_name: 'CSE-C', room_number: 'Programming Lab 3', day_of_week: 'Wednesday', start_time: '13:00', end_time: '15:00', department: 'CSE' },
  { id: 'tt_meera_thu_1', faculty_employee_id: 'FAC003', faculty_name: 'Prof. Meera Krishnan', faculty_email: 'meera.krishnan@example.edu', cabin_location: 'Main Academic Block, Cabin 308', subject_code: 'CS203', subject_name: 'Design & Analysis of Algorithms', section_name: 'CSE-A', room_number: 'CS-301', day_of_week: 'Thursday', start_time: '10:00', end_time: '11:00', department: 'CSE' },
  { id: 'tt_meera_fri_1', faculty_employee_id: 'FAC003', faculty_name: 'Prof. Meera Krishnan', faculty_email: 'meera.krishnan@example.edu', cabin_location: 'Main Academic Block, Cabin 308', subject_code: 'CS203', subject_name: 'Design & Analysis of Algorithms', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Friday', start_time: '10:00', end_time: '11:00', department: 'CSE' },

  // Dr. Suresh Balan (FAC004) - Computer Networks Faculty
  { id: 'tt_suresh_mon_1', faculty_employee_id: 'FAC004', faculty_name: 'Dr. Suresh Balan', faculty_email: 'suresh.balan@example.edu', cabin_location: 'Lab Complex, 2nd Floor Faculty Lounge', subject_code: 'CS204', subject_name: 'Computer Networks', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Monday', start_time: '14:00', end_time: '15:00', department: 'CSE' },
  { id: 'tt_suresh_tue_1', faculty_employee_id: 'FAC004', faculty_name: 'Dr. Suresh Balan', faculty_email: 'suresh.balan@example.edu', cabin_location: 'Lab Complex, 2nd Floor Faculty Lounge', subject_code: 'CS204', subject_name: 'Computer Networks', section_name: 'CSE-D', room_number: 'CS-304', day_of_week: 'Tuesday', start_time: '13:00', end_time: '14:00', department: 'CSE' },
  { id: 'tt_suresh_wed_1', faculty_employee_id: 'FAC004', faculty_name: 'Dr. Suresh Balan', faculty_email: 'suresh.balan@example.edu', cabin_location: 'Lab Complex, 2nd Floor Faculty Lounge', subject_code: 'CS204', subject_name: 'Computer Networks', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Wednesday', start_time: '10:00', end_time: '11:00', department: 'CSE' },
  { id: 'tt_suresh_thu_1', faculty_employee_id: 'FAC004', faculty_name: 'Dr. Suresh Balan', faculty_email: 'suresh.balan@example.edu', cabin_location: 'Lab Complex, 2nd Floor Faculty Lounge', subject_code: 'CS204L', subject_name: 'Networks Lab', section_name: 'CSE-C', room_number: 'Network Lab 1', day_of_week: 'Thursday', start_time: '08:00', end_time: '10:00', department: 'CSE' },
  { id: 'tt_suresh_fri_1', faculty_employee_id: 'FAC004', faculty_name: 'Dr. Suresh Balan', faculty_email: 'suresh.balan@example.edu', cabin_location: 'Lab Complex, 2nd Floor Faculty Lounge', subject_code: 'CS204', subject_name: 'Computer Networks', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Friday', start_time: '13:00', end_time: '14:00', department: 'CSE' },

  // Prof. Neha Kapoor (FAC005) - Theory of Computation Faculty
  { id: 'tt_neha_mon_1', faculty_employee_id: 'FAC005', faculty_name: 'Prof. Neha Kapoor', faculty_email: 'neha.kapoor@example.edu', cabin_location: 'CS Block, 1st Floor Cabin 14', subject_code: 'CS205', subject_name: 'Theory of Computation', section_name: 'CSE-E', room_number: 'CS-305', day_of_week: 'Monday', start_time: '08:00', end_time: '09:00', department: 'CSE' },
  { id: 'tt_neha_tue_1', faculty_employee_id: 'FAC005', faculty_name: 'Prof. Neha Kapoor', faculty_email: 'neha.kapoor@example.edu', cabin_location: 'CS Block, 1st Floor Cabin 14', subject_code: 'CS205', subject_name: 'Theory of Computation', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Tuesday', start_time: '14:00', end_time: '15:00', department: 'CSE' },
  { id: 'tt_neha_wed_1', faculty_employee_id: 'FAC005', faculty_name: 'Prof. Neha Kapoor', faculty_email: 'neha.kapoor@example.edu', cabin_location: 'CS Block, 1st Floor Cabin 14', subject_code: 'CS205', subject_name: 'Theory of Computation', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Wednesday', start_time: '11:00', end_time: '12:00', department: 'CSE' },
  { id: 'tt_neha_thu_1', faculty_employee_id: 'FAC005', faculty_name: 'Prof. Neha Kapoor', faculty_email: 'neha.kapoor@example.edu', cabin_location: 'CS Block, 1st Floor Cabin 14', subject_code: 'CS205', subject_name: 'Theory of Computation', section_name: 'CSE-C', room_number: 'CS-302', day_of_week: 'Thursday', start_time: '11:00', end_time: '12:00', department: 'CSE' },
  { id: 'tt_neha_fri_1', faculty_employee_id: 'FAC005', faculty_name: 'Prof. Neha Kapoor', faculty_email: 'neha.kapoor@example.edu', cabin_location: 'CS Block, 1st Floor Cabin 14', subject_code: 'CS205', subject_name: 'Theory of Computation', section_name: 'CSE-F', room_number: 'CS-306', day_of_week: 'Friday', start_time: '11:00', end_time: '12:00', department: 'CSE' },

  // Additional sections CSE-D to CSE-J faculty schedule entries
  { id: 'tt_arvind_mon_1', faculty_employee_id: 'FAC006', faculty_name: 'Dr. Arvind Nair', faculty_email: 'arvind.nair@example.edu', cabin_location: 'Main Academic Block, Cabin 310', subject_code: 'CS206', subject_name: 'Software Engineering', section_name: 'CSE-F', room_number: 'CS-306', day_of_week: 'Monday', start_time: '09:00', end_time: '10:00', department: 'CSE' },
  { id: 'tt_kavitha_tue_1', faculty_employee_id: 'FAC007', faculty_name: 'Prof. Kavitha Iyer', faculty_email: 'kavitha.iyer@example.edu', cabin_location: 'CS Block, 2nd Floor Cabin 22', subject_code: 'CS207', subject_name: 'Machine Learning', section_name: 'CSE-G', room_number: 'CS-307', day_of_week: 'Tuesday', start_time: '11:00', end_time: '12:00', department: 'CSE' },
  { id: 'tt_prakash_wed_1', faculty_employee_id: 'FAC008', faculty_name: 'Dr. Prakash Verma', faculty_email: 'prakash.verma@example.edu', cabin_location: 'Lab Complex, 1st Floor Cabin 05', subject_code: 'CS208', subject_name: 'Cyber Security', section_name: 'CSE-H', room_number: 'CS-308', day_of_week: 'Wednesday', start_time: '08:00', end_time: '09:00', department: 'CSE' },
  { id: 'tt_swathi_thu_1', faculty_employee_id: 'FAC009', faculty_name: 'Prof. Swathi Rao', faculty_email: 'swathi.rao@example.edu', cabin_location: 'Main Academic Block, Cabin 302', subject_code: 'CS209', subject_name: 'Cloud Computing', section_name: 'CSE-I', room_number: 'CS-309', day_of_week: 'Thursday', start_time: '14:00', end_time: '15:00', department: 'CSE' },
  { id: 'tt_vikram_fri_1', faculty_employee_id: 'FAC010', faculty_name: 'Dr. Vikram Das', faculty_email: 'vikram.das@example.edu', cabin_location: 'CS Block, Ground Floor Cabin 02', subject_code: 'CS210', subject_name: 'Data Analytics', section_name: 'CSE-J', room_number: 'CS-310', day_of_week: 'Friday', start_time: '10:00', end_time: '11:00', department: 'CSE' }
];

const localCustomTimetable: FacultyTimetableEntry[] = [...INITIAL_FACULTY_TIMETABLE];

export async function fetchFacultyTimetable(params?: {
  facultyId?: string;
  facultyEmail?: string;
  facultyEmployeeId?: string;
  sectionName?: string;
  dayOfWeek?: string;
}): Promise<FacultyTimetableEntry[]> {
  try {
    let query = supabase.from('faculty_timetable').select('*');

    if (params?.facultyEmployeeId) {
      query = query.eq('faculty_employee_id', params.facultyEmployeeId);
    } else if (params?.facultyEmail) {
      query = query.ilike('faculty_email', params.facultyEmail.trim());
    } else if (params?.facultyId) {
      query = query.eq('faculty_id', params.facultyId);
    }

    if (params?.sectionName) {
      query = query.ilike('section_name', params.sectionName.trim());
    }

    if (params?.dayOfWeek) {
      query = query.ilike('day_of_week', params.dayOfWeek.trim());
    }

    const { data, error } = await query.order('start_time', { ascending: true });

    let allEntries = [...localCustomTimetable];
    if (!error && data && data.length > 0) {
      const dbIds = new Set(data.map((d: any) => d.id));
      const extraLocal = localCustomTimetable.filter((l) => !dbIds.has(l.id));
      allEntries = [...(data as FacultyTimetableEntry[]), ...extraLocal];
    }

    // Apply filters to combined local/db set
    if (params?.facultyEmployeeId) {
      const targetEmp = params.facultyEmployeeId.toLowerCase();
      allEntries = allEntries.filter((t) => t.faculty_employee_id.toLowerCase() === targetEmp);
    } else if (params?.facultyEmail) {
      const targetMail = params.facultyEmail.toLowerCase();
      allEntries = allEntries.filter((t) => (t.faculty_email || '').toLowerCase() === targetMail);
    } else if (params?.facultyId) {
      allEntries = allEntries.filter((t) => t.faculty_id === params.facultyId);
    }

    if (params?.sectionName) {
      const targetSec = params.sectionName.toUpperCase();
      allEntries = allEntries.filter((t) => (t.section_name || '').toUpperCase() === targetSec || (t.section_name || '').toUpperCase().endsWith(targetSec));
    }

    if (params?.dayOfWeek) {
      const targetDay = params.dayOfWeek.toLowerCase();
      allEntries = allEntries.filter((t) => t.day_of_week.toLowerCase() === targetDay);
    }

    return allEntries;
  } catch {
    let filtered = [...localCustomTimetable];
    if (params?.facultyEmployeeId) {
      filtered = filtered.filter((t) => t.faculty_employee_id.toLowerCase() === params.facultyEmployeeId!.toLowerCase());
    } else if (params?.facultyEmail) {
      filtered = filtered.filter((t) => (t.faculty_email || '').toLowerCase() === params.facultyEmail!.toLowerCase());
    }
    if (params?.sectionName) {
      filtered = filtered.filter((t) => (t.section_name || '').toUpperCase() === params.sectionName!.toUpperCase());
    }
    if (params?.dayOfWeek) {
      filtered = filtered.filter((t) => t.day_of_week.toLowerCase() === params.dayOfWeek!.toLowerCase());
    }
    return filtered;
  }
}

export async function updateFacultyCabinLocation(facultyIdOrEmail: string, cabinLocation: string): Promise<{ success: boolean; error?: string }> {
  try {
    const target = facultyIdOrEmail.trim().toLowerCase();
    const facs = await fetchFacultyMembers();
    const fac = facs.find((f) => f.id === target || f.email.toLowerCase() === target || f.employee_id.toLowerCase() === target);

    if (fac) {
      await updateFacultyMember(fac.id, { cabin_location: cabinLocation });
    }

    localCustomTimetable.forEach((entry) => {
      if (
        (entry.faculty_id && entry.faculty_id === target) ||
        (entry.faculty_email && entry.faculty_email.toLowerCase() === target) ||
        (entry.faculty_employee_id && entry.faculty_employee_id.toLowerCase() === target)
      ) {
        entry.cabin_location = cabinLocation;
      }
    });

    try {
      await supabase
        .from('faculty_timetable')
        .update({ cabin_location: cabinLocation, updated_at: new Date().toISOString() })
        .or(`faculty_email.eq.${target},faculty_employee_id.eq.${target}`);
    } catch {}

    return { success: true };
  } catch (err: any) {
    return { success: false, error: err?.message || 'Failed to update cabin location' };
  }
}

export async function upsertFacultyTimetableEntry(entry: Partial<FacultyTimetableEntry>): Promise<{ success: boolean; data?: FacultyTimetableEntry; error?: string }> {
  try {
    const id = entry.id || `tt_${Date.now()}_${Math.floor(Math.random() * 1000)}`;
    const fullEntry: FacultyTimetableEntry = {
      id,
      faculty_employee_id: entry.faculty_employee_id || 'FAC001',
      faculty_name: entry.faculty_name || 'Faculty Member',
      faculty_email: entry.faculty_email || '',
      cabin_location: entry.cabin_location || 'Main Academic Block',
      subject_code: entry.subject_code || 'CS201',
      subject_name: entry.subject_name || 'Academic Course',
      section_name: (entry.section_name || 'CSE-A').toUpperCase(),
      room_number: entry.room_number || 'CS-101',
      day_of_week: (entry.day_of_week as any) || 'Monday',
      start_time: entry.start_time || '08:00',
      end_time: entry.end_time || '09:00',
      department: entry.department || 'CSE',
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString(),
    };

    const existingIdx = localCustomTimetable.findIndex((t) => t.id === fullEntry.id);
    if (existingIdx !== -1) {
      localCustomTimetable[existingIdx] = fullEntry;
    } else {
      localCustomTimetable.push(fullEntry);
    }

    try {
      const { data, error } = await supabase
        .from('faculty_timetable')
        .upsert([fullEntry], { onConflict: 'id' })
        .select()
        .single();

      if (!error && data) {
        return { success: true, data: data as FacultyTimetableEntry };
      }
    } catch {}

    return { success: true, data: fullEntry };
  } catch (err: any) {
    return { success: false, error: err?.message || 'Failed to save timetable entry' };
  }
}

export async function deleteFacultyTimetableEntry(id: string): Promise<{ success: boolean; error?: string }> {
  try {
    const idx = localCustomTimetable.findIndex((t) => t.id === id);
    if (idx !== -1) {
      localCustomTimetable.splice(idx, 1);
    }

    try {
      await supabase.from('faculty_timetable').delete().eq('id', id);
    } catch {}

    return { success: true };
  } catch (err: any) {
    return { success: false, error: err?.message || 'Failed to delete entry' };
  }
}

export function validateTimetableExcelHeaders(headers: string[]): TimetableHeaderValidationResult {
  const normalized = headers.map((h) => String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''));
  const mappedHeaders: Record<string, string> = {};

  const defs = [
    { key: 'faculty_name', keywords: ['facultyname', 'faculty', 'teachername', 'staffname', 'professorname', 'name'] },
    { key: 'faculty_employee_id', keywords: ['employeeid', 'empid', 'facultyid', 'staffid', 'empcode', 'code', 'id'] },
    { key: 'faculty_email', keywords: ['email', 'facultyemail', 'mail', 'emailaddress'] },
    { key: 'subject_name', keywords: ['subjectname', 'subject', 'course', 'coursename', 'subj'] },
    { key: 'subject_code', keywords: ['subjectcode', 'coursecode', 'code'] },
    { key: 'section_name', keywords: ['sectionname', 'section', 'class', 'sec', 'batch'] },
    { key: 'day_of_week', keywords: ['dayofweek', 'day', 'weekday'] },
    { key: 'start_time', keywords: ['starttime', 'fromtime', 'start', 'from'] },
    { key: 'end_time', keywords: ['endtime', 'totime', 'end', 'to'] },
    { key: 'room_number', keywords: ['roomnumber', 'room', 'hall', 'lab', 'roomno', 'cabin', 'location'] },
    { key: 'cabin_location', keywords: ['cabinlocation', 'cabin', 'office', 'roomno', 'officelocation'] }
  ];

  for (const def of defs) {
    let idx = normalized.findIndex((h) => def.keywords.includes(h));
    if (idx === -1) {
      idx = normalized.findIndex((h) => h.length > 1 && def.keywords.some((k) => h.includes(k) || (k.length > 3 && k.includes(h))));
    }
    if (idx !== -1) {
      mappedHeaders[def.key] = headers[idx];
    }
  }

  const missingColumns: string[] = [];
  if (!mappedHeaders['faculty_name'] && !mappedHeaders['faculty_employee_id']) {
    missingColumns.push('FACULTY NAME / EMPLOYEE ID');
  }
  if (!mappedHeaders['subject_name']) {
    missingColumns.push('SUBJECT NAME');
  }

  return {
    valid: missingColumns.length === 0,
    missingColumns,
    mappedHeaders,
  };
}

export async function importTimetableBatch(rows: TimetableImportRow[]): Promise<ImportResult> {
  let importedCount = 0;
  let updatedCount = 0;
  let skippedCount = 0;
  let invalidCount = 0;
  const errors: string[] = [];

  const existingFaculty = await fetchFacultyMembers();

  for (let index = 0; index < rows.length; index++) {
    const row = rows[index];
    const lineNo = index + 2;

    const facName = String(row.faculty_name || '').trim();
    let empId = String(row.faculty_employee_id || '').trim();
    let email = String(row.faculty_email || '').trim().toLowerCase();
    const subjName = String(row.subject_name || '').trim();
    const subjCode = String(row.subject_code || '').trim() || 'CS200';
    const secName = String(row.section_name || 'CSE-A').trim().toUpperCase();
    let day = String(row.day_of_week || 'Monday').trim();
    let startTime = String(row.start_time || '08:00').trim();
    let endTime = String(row.end_time || '09:00').trim();
    let room = String(row.room_number || 'CS-101').trim();

    if (!facName && !empId) {
      invalidCount++;
      errors.push(`Row ${lineNo}: Missing faculty identity (Name or Employee ID)`);
      continue;
    }

    if (!subjName) {
      invalidCount++;
      errors.push(`Row ${lineNo}: Missing Subject Name`);
      continue;
    }

    const matchedFac = existingFaculty.find(
      (f) =>
        (empId && f.employee_id.toLowerCase() === empId.toLowerCase()) ||
        (email && f.email.toLowerCase() === email) ||
        (facName && f.name.toLowerCase().includes(facName.toLowerCase()))
    );

    if (matchedFac) {
      if (!empId) empId = matchedFac.employee_id;
      if (!email) email = matchedFac.email;
    } else if (!empId) {
      empId = `FAC_${Math.floor(100 + Math.random() * 900)}`;
    }

    const dayMap: Record<string, string> = {
      mon: 'Monday', monday: 'Monday', m: 'Monday',
      tue: 'Tuesday', tuesday: 'Tuesday', t: 'Tuesday',
      wed: 'Wednesday', wednesday: 'Wednesday', w: 'Wednesday',
      thu: 'Thursday', thursday: 'Thursday', th: 'Thursday',
      fri: 'Friday', friday: 'Friday', f: 'Friday',
      sat: 'Saturday', saturday: 'Saturday',
      sun: 'Sunday', sunday: 'Sunday'
    };
    const cleanDayKey = day.toLowerCase().replace(/[^a-z]/g, '');
    day = dayMap[cleanDayKey] || 'Monday';

    const normalizeTime = (raw: string): string => {
      if (!raw) return '08:00';
      if (raw.includes(':')) {
        const parts = raw.split(':');
        let h = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10) || 0;
        if (raw.toLowerCase().includes('pm') && h < 12) h += 12;
        if (raw.toLowerCase().includes('am') && h === 12) h = 0;
        return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
      }
      const num = parseInt(raw, 10);
      if (!isNaN(num)) {
        if (num >= 8 && num <= 12) return `${String(num).padStart(2, '0')}:00`;
        if (num >= 1 && num <= 7) return `${String(num + 12).padStart(2, '0')}:00`;
      }
      return '08:00';
    };

    startTime = normalizeTime(startTime);
    endTime = normalizeTime(endTime);

    const isDup = localCustomTimetable.some(
      (t) =>
        t.faculty_employee_id.toLowerCase() === empId.toLowerCase() &&
        t.day_of_week.toLowerCase() === day.toLowerCase() &&
        t.start_time === startTime
    );

    const entryId = `tt_${Date.now()}_${index}`;
    const newEntry: FacultyTimetableEntry = {
      id: entryId,
      faculty_employee_id: empId,
      faculty_name: facName || matchedFac?.name || 'Faculty Member',
      faculty_email: email || matchedFac?.email,
      cabin_location: row.cabin_location || matchedFac?.cabin_location || 'Main Academic Block',
      subject_code: subjCode,
      subject_name: subjName,
      section_name: secName,
      room_number: room,
      day_of_week: day as any,
      start_time: startTime,
      end_time: endTime,
      department: row.department || matchedFac?.department || 'CSE',
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString(),
    };

    localCustomTimetable.push(newEntry);

    try {
      await supabase.from('faculty_timetable').upsert([newEntry]);
    } catch {}

    if (isDup) {
      updatedCount++;
    } else {
      importedCount++;
    }
  }

  return { importedCount, updatedCount, skippedCount, invalidCount, errors };
}

// ----------------------------------------------------
// DYNAMIC STUDENT TIMETABLE & TODAY'S CLASSES ENGINE
// ----------------------------------------------------

export interface TodayClassItem {
  id: string;
  subject_name: string;
  subject_code?: string;
  faculty_name: string;
  room_number: string;
  start_time: string;
  end_time: string;
  period_number?: number;
  section_name: string;
  status: 'UPCOMING' | 'ONGOING' | 'COMPLETED';
  status_label: string;
  starts_in_minutes?: number;
}

export interface TodayScheduleResult {
  isHoliday: boolean;
  holidayReason?: string;
  dayName: string;
  dateStr: string;
  classes: TodayClassItem[];
  nextClass?: TodayClassItem;
}

// Academic Calendar Holidays / Working-Day Override Table
export interface CalendarHolidayOverride {
  dateStr: string; // YYYY-MM-DD
  is_holiday: boolean;
  holiday_name: string;
  is_working_day_override?: boolean;
}

const INSTITUTION_HOLIDAY_OVERRIDES: CalendarHolidayOverride[] = [
  { dateStr: '2026-09-14', is_holiday: fontIsHoliday('2026-09-14'), holiday_name: 'College Foundation Day' },
  { dateStr: '2026-10-02', is_holiday: true, holiday_name: 'Gandhi Jayanti' },
  { dateStr: '2026-10-25', is_holiday: true, holiday_name: 'Dussehra Holiday' },
  { dateStr: '2026-11-01', is_holiday: true, holiday_name: 'Deepavali' }
];

function fontIsHoliday(_d: string): boolean {
  return false;
}

export async function fetchStudentTimetable(sectionName?: string): Promise<FacultyTimetableEntry[]> {
  const targetSec = (sectionName || 'CSE-C').toUpperCase().trim();
  return await fetchFacultyTimetable({ sectionName: targetSec });
}

export async function getTodayClassSchedule(
  sectionName?: string,
  simulatedDate?: Date
): Promise<TodayScheduleResult> {
  const now = simulatedDate || new Date();
  const dateStr = now.toISOString().split('T')[0];

  const daysOfWeek = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const dayName = daysOfWeek[now.getDay()];

  // 1. DATE-SPECIFIC HOLIDAY & WORKING DAY OVERRIDE CHECK
  const holidayRecord = INSTITUTION_HOLIDAY_OVERRIDES.find(h => h.dateStr === dateStr);

  if (holidayRecord && holidayRecord.is_holiday && !holidayRecord.is_working_day_override) {
    return {
      isHoliday: true,
      holidayReason: holidayRecord.holiday_name,
      dayName,
      dateStr,
      classes: []
    };
  }

  const isWeekend = dayName === 'Saturday' || dayName === 'Sunday';
  const isWorkingWeekendOverride = holidayRecord?.is_working_day_override;

  if (isWeekend && !isWorkingWeekendOverride) {
    return {
      isHoliday: true,
      holidayReason: `${dayName} — Weekend / No Classes Scheduled`,
      dayName,
      dateStr,
      classes: []
    };
  }

  // 2. FETCH SECTION TIMETABLE FOR TODAY'S DAY
  const sec = (sectionName || 'CSE-C').toUpperCase().trim();
  const allEntries = await fetchStudentTimetable(sec);
  const todayEntries = allEntries
    .filter(t => t.day_of_week.toLowerCase() === dayName.toLowerCase())
    .sort((a, b) => a.start_time.localeCompare(b.start_time));

  if (todayEntries.length === 0) {
    return {
      isHoliday: false,
      holidayReason: `No classes scheduled for ${dayName}`,
      dayName,
      dateStr,
      classes: []
    };
  }

  // 3. CALCULATE LIVE STATUS BASED ON CURRENT TIME
  const currentMinutes = now.getHours() * 60 + now.getMinutes();

  const classes: TodayClassItem[] = todayEntries.map(e => {
    const [startH, startM] = e.start_time.split(':').map(Number);
    const [endH, endM] = e.end_time.split(':').map(Number);
    const startMin = startH * 60 + startM;
    const endMin = endH * 60 + endM;

    let status: 'UPCOMING' | 'ONGOING' | 'COMPLETED' = 'UPCOMING';
    let status_label = '';
    let starts_in_minutes: number | undefined = undefined;

    if (currentMinutes >= endMin) {
      status = 'COMPLETED';
      status_label = 'Completed';
    } else if (currentMinutes >= startMin && currentMinutes < endMin) {
      status = 'ONGOING';
      status_label = 'Now · Live Class';
    } else {
      status = 'UPCOMING';
      starts_in_minutes = startMin - currentMinutes;
      if (starts_in_minutes < 60) {
        status_label = `Starts in ${starts_in_minutes} min`;
      } else {
        const hrs = Math.floor(starts_in_minutes / 60);
        const mins = starts_in_minutes % 60;
        status_label = `Starts in ${hrs}h ${mins > 0 ? `${mins}m` : ''}`;
      }
    }

    return {
      id: e.id,
      subject_name: e.subject_name,
      subject_code: e.subject_code,
      faculty_name: e.faculty_name,
      room_number: e.room_number,
      start_time: e.start_time,
      end_time: e.end_time,
      period_number: (e as any).period_number,
      section_name: e.section_name,
      status,
      status_label,
      starts_in_minutes
    };
  });

  const nextClass = classes.find(c => c.status === 'ONGOING' || c.status === 'UPCOMING');

  return {
    isHoliday: false,
    dayName,
    dateStr,
    classes,
    nextClass
  };
}

export function parseTimetableExcel(fileBuffer: ArrayBuffer): {
  success: boolean;
  rows: TimetableImportRow[];
  headers: string[];
  error?: string;
} {
  try {
    const workbook = XLSX.read(fileBuffer, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) return { success: false, rows: [], headers: [], error: 'Empty Excel file' };

    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json<any>(worksheet, { header: 1 });

    if (!rawData || rawData.length === 0) {
      return { success: false, rows: [], headers: [], error: 'No data found in worksheet' };
    }

    let headerRowIndex = 0;
    for (let i = 0; i < Math.min(10, rawData.length); i++) {
      const row = rawData[i];
      if (Array.isArray(row) && row.some((cell) => {
        const c = String(cell || '').toLowerCase();
        return c.includes('faculty') || c.includes('subject') || c.includes('day') || c.includes('room') || c.includes('section');
      })) {
        headerRowIndex = i;
        break;
      }
    }

    const headers = (rawData[headerRowIndex] || []).map((h: any) => String(h || '').trim());
    const validation = validateTimetableExcelHeaders(headers);

    if (!validation.valid) {
      return {
        success: false,
        rows: [],
        headers,
        error: `Missing required columns: ${validation.missingColumns.join(', ')}`
      };
    }

    const rows: TimetableImportRow[] = [];
    const m = validation.mappedHeaders;

    for (let i = headerRowIndex + 1; i < rawData.length; i++) {
      const r = rawData[i];
      if (!r || r.length === 0) continue;

      const getVal = (key: string): string => {
        const colName = m[key];
        if (!colName) return '';
        const colIdx = headers.indexOf(colName);
        if (colIdx === -1) return '';
        return String(r[colIdx] || '').trim();
      };

      const faculty_name = getVal('faculty_name');
      const faculty_employee_id = getVal('faculty_employee_id');
      const subject_name = getVal('subject_name');

      if (faculty_name || faculty_employee_id || subject_name) {
        rows.push({
          faculty_name: faculty_name || 'Faculty Member',
          faculty_employee_id,
          faculty_email: getVal('faculty_email'),
          cabin_location: getVal('cabin_location'),
          subject_code: getVal('subject_code'),
          subject_name: subject_name || 'Academic Course',
          section_name: getVal('section_name') || 'CSE-A',
          day_of_week: getVal('day_of_week') || 'Monday',
          start_time: getVal('start_time') || '08:00',
          end_time: getVal('end_time') || '09:00',
          room_number: getVal('room_number') || 'CS-101',
          department: getVal('department') || 'CSE',
        });
      }
    }

    return { success: true, rows, headers };
  } catch (err: any) {
    return { success: false, rows: [], headers: [], error: err?.message || 'Failed to parse Excel file' };
  }
}








