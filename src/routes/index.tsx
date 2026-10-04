import { createFileRoute } from '@tanstack/react-router';
import { useEffect, useRef, useState } from 'react';
import { BookOpen, CalendarDays, ChartNoAxesColumn, ChevronRight, CircleHelp, ClipboardList, Download, FileSpreadsheet, FileText, GraduationCap, LayoutDashboard, LogOut, Menu, Plus, Search, ShieldCheck, UploadCloud, Users, X } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { supabase } from '@/integrations/supabase/client';
import { lovable } from '@/integrations/lovable/index';
import type { Database } from '@/integrations/supabase/types';
import * as XLSX from 'xlsx';
import campus from '@/assets/campus.jpg';

type Schedule = Database['public']['Tables']['schedules']['Row'];
type Student = Database['public']['Tables']['students']['Row'];
type Grade = Database['public']['Tables']['grades']['Row'];
type Upload = Database['public']['Tables']['grade_uploads']['Row'];
type Syllabus = Database['public']['Tables']['syllabus']['Row'];
type Role = Database['public']['Enums']['app_role'];
type Section = 'overview' | 'schedule' | 'reports' | 'syllabus';
const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const sampleSchedules: Schedule[] = [
  { id:'s1', subject:'Mathematics 10', section_name:'Grade 10 - St. Anne', teacher_name:'Ms. R. Santos', room_name:'Room 204', day_of_week:1, start_time:'08:00', end_time:'09:00' },
  { id:'s2', subject:'English 10', section_name:'Grade 10 - St. Anne', teacher_name:'Mr. D. Reyes', room_name:'Room 204', day_of_week:1, start_time:'09:15', end_time:'10:15' },
  { id:'s3', subject:'Science 10', section_name:'Grade 10 - St. Anne', teacher_name:'Ms. A. Cruz', room_name:'Lab 1', day_of_week:1, start_time:'10:30', end_time:'11:30' },
  { id:'s4', subject:'Filipino 10', section_name:'Grade 10 - St. Anne', teacher_name:'Mrs. L. Garcia', room_name:'Room 204', day_of_week:2, start_time:'08:00', end_time:'09:00' },
  { id:'s5', subject:'Araling Panlipunan', section_name:'Grade 10 - St. Anne', teacher_name:'Mr. M. Dela Cruz', room_name:'Room 204', day_of_week:3, start_time:'09:15', end_time:'10:15' },
];
const sampleSyllabus: Syllabus[] = [
  { id:'a', subject:'Mathematics 10', section_name:'Grade 10 - St. Anne', teacher_name:'Ms. R. Santos', overview:'Explore patterns, functions, and the mathematics of everyday life.', topics:'Quadratic equations, Polynomials, Geometry, Probability', updated_at:'' },
  { id:'b', subject:'English 10', section_name:'Grade 10 - St. Anne', teacher_name:'Mr. D. Reyes', overview:'Develop critical reading, clear writing, and confident communication.', topics:'World literature, Writing, Speech, Research', updated_at:'' },
  { id:'c', subject:'Science 10', section_name:'Grade 10 - St. Anne', teacher_name:'Ms. A. Cruz', overview:'Understand the natural world through inquiry and investigation.', topics:'Earth science, Chemistry, Physics, Environment', updated_at:'' },
  { id:'d', subject:'Filipino 10', section_name:'Grade 10 - St. Anne', teacher_name:'Mrs. L. Garcia', overview:'Paglinang ng wika, panitikan, at kulturang Pilipino.', topics:'Panitikan, Wika, Pagsulat, Pagsusuri', updated_at:'' },
];
const sampleStudents: Student[] = [
  { id:'p1', full_name:'Andrea M. Reyes', student_number:'SCCV-2026-001', section_name:'Grade 10 - St. Anne', user_id:null, parent_id:null, created_at:'' },
  { id:'p2', full_name:'Joshua P. Santos', student_number:'SCCV-2026-002', section_name:'Grade 10 - St. Anne', user_id:null, parent_id:null, created_at:'' },
  { id:'p3', full_name:'Sofia L. Cruz', student_number:'SCCV-2026-003', section_name:'Grade 10 - St. Anne', user_id:null, parent_id:null, created_at:'' },
];
const sampleUploads: Upload[] = [{id:'u1',file_name:'math10_grades.xlsx',subject:'Mathematics 10',section_name:'Grade 10 - St. Anne',uploaded_by:'',uploaded_at:'2026-10-02',status:'approved'}];
const sampleGrades: Grade[] = [
  {id:'g1',upload_id:'u1',student_id:'p1',subject:'Mathematics 10',prelim:89,midterm:92,final:94,computed_grade:91.67},
  {id:'g2',upload_id:'u1',student_id:'p2',subject:'Mathematics 10',prelim:86,midterm:88,final:90,computed_grade:88},
  {id:'g3',upload_id:'u1',student_id:'p3',subject:'Mathematics 10',prelim:93,midterm:95,final:96,computed_grade:94.67},
];
const nav: {id:Section; label:string; icon:typeof LayoutDashboard}[] = [
  {id:'overview',label:'Overview',icon:LayoutDashboard}, {id:'schedule',label:'Schedule',icon:CalendarDays},
  {id:'reports',label:'Reports & grades',icon:ChartNoAxesColumn}, {id:'syllabus',label:'Syllabus',icon:BookOpen},
];
const labelTime = (time:string) => { const [h,m] = time.split(':').map(Number); return `${h % 12 || 12}:${String(m).padStart(2,'0')} ${h >= 12 ? 'PM':'AM'}`; };
const errorText = (e:unknown) => e instanceof Error ? e.message : String(e);

export const Route = createFileRoute('/')({
  head: () => ({ meta: [
    {title:'Academic Portal | St. Catherine College of Valenzuela'},
    {name:'description',content:'Schedules, reports, syllabus, and approved grades for the St. Catherine College of Valenzuela community.'},
    {property:'og:title',content:'Academic Portal | St. Catherine College of Valenzuela'},
    {property:'og:description',content:'View schedules, syllabus, and approved academic reports in one place.'},
    {property:'og:type',content:'website'}, {name:'twitter:card',content:'summary_large_image'},
  ]}),
  component: Portal,
});

function Portal() {
  const [section,setSection] = useState<Section>('overview');
  const [day,setDay] = useState(1);
  const [session,setSession] = useState<Awaited<ReturnType<typeof supabase.auth.getSession>>['data']['session']>(null);
  const [role,setRole] = useState<Role|null>(null);
  const [name,setName] = useState('');
  const [schedules,setSchedules] = useState<Schedule[]>([]);
  const [syllabus,setSyllabus] = useState<Syllabus[]>([]);
  const [students,setStudents] = useState<Student[]>([]);
  const [uploads,setUploads] = useState<Upload[]>([]);
  const [grades,setGrades] = useState<Grade[]>([]);
  const [requests,setRequests] = useState<Database['public']['Tables']['document_requests']['Row'][]>([]);
  const [modal,setModal] = useState<'auth'|'upload'|'schedule'|'syllabus'|'student'|'request'|null>(null);
  const [authMode,setAuthMode] = useState<'sign-in'|'sign-up'>('sign-in');
  const [notice,setNotice] = useState('');
  const [busy,setBusy] = useState(false);
  const [search,setSearch] = useState('');
  const [period,setPeriod] = useState('All');
  const [mobileMenu,setMobileMenu] = useState(false);
  const fileRef = useRef<HTMLInputElement>(null);
  const userId = session?.user.id;
  const isStaff = role === 'admin' || role === 'teacher' || role === 'adviser';
  const canUpload = role === 'admin' || role === 'teacher';
  const demo = !session;
  const displaySchedules = demo ? sampleSchedules : schedules;
  const displaySyllabus = demo ? sampleSyllabus : syllabus;
  const displayStudents = demo ? sampleStudents : students;
  const displayUploads = demo ? sampleUploads : uploads;
  const displayGrades = demo ? sampleGrades : grades;
  const visibleGrades = displayGrades.filter(g => demo || isStaff || displayUploads.some(u => u.id === g.upload_id && u.status === 'approved'));

  async function refresh() {
    const [a,b,c,d,e,f,g] = await Promise.all([
      supabase.from('schedules').select('*').order('start_time'), supabase.from('syllabus').select('*').order('subject'),
      supabase.from('students').select('*').order('full_name'), supabase.from('grade_uploads').select('*').order('uploaded_at',{ascending:false}),
      supabase.from('grades').select('*'),supabase.from('user_roles').select('role'),supabase.from('document_requests').select('*').order('requested_at',{ascending:false}),
    ]);
    setSchedules(a.data ?? []); setSyllabus(b.data ?? []); setStudents(c.data ?? []); setUploads(d.data ?? []); setGrades(e.data ?? []); setRole(f.data?.[0]?.role ?? null); setRequests(g.data ?? []);
    const {data:{user}} = await supabase.auth.getUser();
    if(user) { const {data} = await supabase.from('profiles').select('full_name').eq('id',user.id).maybeSingle(); setName(data?.full_name || user.email?.split('@')[0] || 'Member'); }
  }
  useEffect(() => {
    supabase.auth.getSession().then(({data}) => { setSession(data.session); if(data.session) void refresh(); });
    const {data:{subscription}} = supabase.auth.onAuthStateChange((_event,next) => { setSession(next); if(next) void refresh(); else {setRole(null);setName('');} });
    return () => subscription.unsubscribe();
  },[]);
  function go(next:Section) { setSection(next); setMobileMenu(false); setNotice(''); }
  function open(next:typeof modal) { if(demo && next !== 'auth') { setModal('auth'); return; } setNotice(''); setModal(next); }
  function close() { setModal(null); setNotice(''); }
  async function authSubmit(e:React.FormEvent<HTMLFormElement>) {
    e.preventDefault(); setBusy(true); setNotice('');
    const form = new FormData(e.currentTarget);
    const email = String(form.get('email') ?? ''); const password = String(form.get('password') ?? '');
    if(authMode === 'sign-up') {
      const {data,error} = await supabase.auth.signUp({email,password});
      if(error) setNotice(error.message);
      else if(data.user) { if(data.session) await supabase.from('profiles').upsert({id:data.user.id,full_name:String(form.get('name') ?? '')}); setNotice('Account created. Check your email to confirm your address before signing in.'); }
    } else { const {error} = await supabase.auth.signInWithPassword({email,password}); if(error) setNotice(error.message); else close(); }
    setBusy(false);
  }
  async function googleSignIn() {
    setBusy(true); const result = await lovable.auth.signInWithOAuth('google',{redirect_uri:window.location.origin});
    if(result.error) setNotice(result.error.message); else if(!result.redirected) close(); setBusy(false);
  }
  async function signOut() { await supabase.auth.signOut(); go('overview'); }
  async function uploadExcel(e:React.FormEvent<HTMLFormElement>) {
    e.preventDefault(); if(!userId || !canUpload) return;
    const form = new FormData(e.currentTarget); const file = fileRef.current?.files?.[0];
    if(!file) {setNotice('Choose an Excel file first.');return;}
    setBusy(true); setNotice('');
    try {
      const book = XLSX.read(await file.arrayBuffer(),{type:'array'});
      const sheet = book.Sheets[book.SheetNames[0]];
      if(!sheet) throw new Error('The workbook has no sheet.');
      const raw = XLSX.utils.sheet_to_json<Record<string,unknown>>(sheet,{defval:''});
      if(!raw.length) throw new Error('The sheet is empty.');
      const keys = Object.keys(raw[0]);
      const get = (row:Record<string,unknown>,names:string[]) => { const key = keys.find(k => names.includes(k.trim().toLowerCase().replace(/[\s_-]+/g,''))); return key ? row[key] : undefined; };
      const parsed = raw.map((row,i) => {
        const number = String(get(row,['studentnumber','studentid','lrn']) ?? '').trim();
        const prelim = Number(get(row,['prelim','preliminary'])); const midterm = Number(get(row,['midterm','midterms'])); const final = Number(get(row,['final','finals']));
        if(!number || ![prelim,midterm,final].every(n => Number.isFinite(n) && n >= 0 && n <= 100) || [prelim,midterm,final].some((_,j) => String(get(row,[['prelim','preliminary'],['midterm','midterms'],['final','finals']][j]) ?? '').trim() === '')) throw new Error(`Row ${i+2}: enter a student number and scores from 0 to 100 in Prelim, Midterm, and Final.`);
        return {number,prelim,midterm,final};
      });
      if(new Set(parsed.map(r=>r.number)).size !== parsed.length) throw new Error('Student numbers must be unique in the file.');
      const subject = String(form.get('subject') ?? '').trim(); const sectionName = String(form.get('section') ?? '').trim();
      const {data:matched,error:studentError} = await supabase.from('students').select('*').in('student_number',parsed.map(r=>r.number));
      if(studentError) throw studentError;
      const valid = (matched ?? []).filter(s => s.section_name === sectionName);
      const missing = parsed.filter(r=>!valid.some(s=>s.student_number === r.number));
      if(missing.length) throw new Error(`Student ${missing.map(r=>r.number).join(', ')} is not registered in ${sectionName}. Add students first.`);
      const {data:upload,error:uploadError} = await supabase.from('grade_uploads').insert({file_name:file.name,subject,section_name:sectionName,uploaded_by:userId}).select().single();
      if(uploadError) throw uploadError;
      const {error:gradeError} = await supabase.from('grades').insert(parsed.map(r=>({upload_id:upload.id,student_id:valid.find(s=>s.student_number===r.number)?.id ?? '',subject,prelim:r.prelim,midterm:r.midterm,final:r.final})));
      if(gradeError) throw gradeError;
      await refresh(); close(); setSection('reports'); setNotice(`${parsed.length} grades submitted for adviser approval.`);
    } catch(error) {setNotice(errorText(error));} finally {setBusy(false);}
  }
  async function approve(upload:Upload) {
    setBusy(true); const {error} = await supabase.from('grade_uploads').update({status:'approved'}).eq('id',upload.id);
    setNotice(error ? error.message : 'Grades approved and now visible to linked students and parents.'); if(!error) await refresh(); setBusy(false);
  }
  async function addSchedule(e:React.FormEvent<HTMLFormElement>) {
    e.preventDefault(); setBusy(true); const f=new FormData(e.currentTarget);
    const {error}=await supabase.from('schedules').insert({subject:String(f.get('subject')),section_name:String(f.get('section')),teacher_name:String(f.get('teacher')),room_name:String(f.get('room')),day_of_week:Number(f.get('day')),start_time:String(f.get('start')),end_time:String(f.get('end'))});
    if(error) setNotice(error.message); else {await refresh();close();setNotice('Class added to the schedule.');} setBusy(false);
  }
  async function addSyllabus(e:React.FormEvent<HTMLFormElement>) {
    e.preventDefault();setBusy(true);const f=new FormData(e.currentTarget);
    const {error}=await supabase.from('syllabus').insert({subject:String(f.get('subject')),section_name:String(f.get('section')),teacher_name:String(f.get('teacher')),overview:String(f.get('overview')),topics:String(f.get('topics'))});
    if(error) setNotice(error.message); else {await refresh();close();setNotice('Syllabus added.');} setBusy(false);
  }
  async function addStudent(e:React.FormEvent<HTMLFormElement>) {
    e.preventDefault();setBusy(true);const f=new FormData(e.currentTarget);
    const {error}=await supabase.from('students').insert({student_number:String(f.get('number')),full_name:String(f.get('name')),section_name:String(f.get('section'))});
    if(error) setNotice(error.message); else {await refresh();close();setNotice('Student registered.');} setBusy(false);
  }
  async function requestDocument(e:React.FormEvent<HTMLFormElement>) {
    e.preventDefault();if(!userId)return;setBusy(true);const f=new FormData(e.currentTarget);
    const {error}=await supabase.from('document_requests').insert({student_id:String(f.get('student')),requested_by:userId,document_type:String(f.get('type'))});
    if(error)setNotice(error.message);else{await refresh();close();setNotice('Document request submitted.');}setBusy(false);
  }
  function downloadReport() {
    const rows=visibleGrades.map(g=>({Student:displayStudents.find(s=>s.id===g.student_id)?.full_name ?? '', 'Student number':displayStudents.find(s=>s.id===g.student_id)?.student_number ?? '',Subject:g.subject,Prelim:g.prelim,Midterm:g.midterm,Final:g.final,Average:g.computed_grade,Status:displayUploads.find(u=>u.id===g.upload_id)?.status ?? ''}));
    const csv=XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(rows));const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));a.download='academic-report.csv';a.click();URL.revokeObjectURL(a.href);
  }
  const title = {overview:'Academic overview',schedule:'Class schedule',reports:'Reports & grades',syllabus:'Course syllabus'}[section];
  return <div className="app-shell">
    <aside className="sidebar">
      <div className="brand"><div className="brand-mark">SC</div><div><div className="brand-name">St. Catherine College</div><div className="brand-sub">of Valenzuela</div></div></div>
      <div className="nav-caption">Workspace</div><div className="side-nav">{nav.map(item=><Button key={item.id} variant="ghost" className={`nav-item ${section===item.id?'active':''}`} onClick={()=>go(item.id)}><item.icon size={17}/>{item.label}</Button>)}</div>
      <div className="sidebar-bottom"><div className="text-xs font-semibold">Academic Portal</div><div className="school-year">School year 2026–2027</div><Button variant="ghost" className="nav-item mt-5 !px-0" onClick={()=>session?void signOut():open('auth')}>{session?<LogOut size={16}/>:<Users size={16}/>} {session?'Sign out':'Sign in'}</Button></div>
    </aside>
    <main className="main-area">
      <header className="topbar"><div className="flex items-center gap-3"><Button variant="ghost" size="icon" className="md:hidden" aria-label="Open menu" onClick={()=>setMobileMenu(!mobileMenu)}><Menu/></Button><div className="topbar-label">Academic Portal</div></div><div className="topbar-actions"><span className="today">School Year 2026–2027</span><span className="avatar-circle">{session?(name.slice(0,2).toUpperCase()):'SC'}</span><div className="hidden sm:block"><div className="text-xs font-bold">{session?name:'Guest preview'}</div><div className="text-[10px] text-muted-foreground capitalize">{session?(role ?? 'Member'):'Sample data'}</div></div></div></header>
      {mobileMenu&&<div className="border-b bg-card p-3 md:hidden">{nav.map(item=><Button key={item.id} variant="ghost" className="w-full justify-start" onClick={()=>go(item.id)}>{item.label}</Button>)}<Button variant="ghost" className="w-full justify-start" onClick={()=>session?void signOut():open('auth')}>{session?'Sign out':'Sign in'}</Button></div>}
      <div className="content">
        <div className="eyebrow">St. Catherine College of Valenzuela</div><h1 className="page-title">{title}</h1><p className="page-intro">{section==='overview'?'Your academic day, all in one place.':section==='schedule'?'A clear view of classes, rooms, and teaching assignments.':section==='reports'?'Review computed grades and share approved results with families.':'Explore subjects and what students will learn this term.'}</p>
        {demo&&<div className="inline-alert flex items-center justify-between gap-3"><span><strong>Preview mode</strong> — showing sample records. Sign in to access school data.</span><Button size="sm" onClick={()=>open('auth')}>Sign in</Button></div>}
        {notice&&!modal&&<div className="inline-alert" role="status">{notice}</div>}
        {section==='overview'&&<>
          <div className="overview-grid"><div className="welcome"><img src={campus} width={1536} height={864} alt="Green college courtyard and covered walkway"/><div className="welcome-content"><span className="welcome-kicker">School year 2026–2027</span><h2>Welcome to your academic space.</h2><p>Stay close to the moments that matter — classes, learning plans, and progress, together in one place.</p></div></div><div className="panel panel-pad"><div className="panel-heading">Quick access</div><p className="panel-sub mt-1">Pick up where you need to be.</p><div className="quick-list"><Button variant="ghost" className="quick-link" onClick={()=>go('schedule')}><span><CalendarDays size={17}/> Class schedule</span><ChevronRight size={15}/></Button><Button variant="ghost" className="quick-link" onClick={()=>go('reports')}><span><ChartNoAxesColumn size={17}/> Reports & grades</span><ChevronRight size={15}/></Button><Button variant="ghost" className="quick-link" onClick={()=>go('syllabus')}><span><BookOpen size={17}/> Course syllabus</span><ChevronRight size={15}/></Button></div></div></div>
          <div className="stat-grid"><div className="panel stat"><div className="stat-top"><div className="stat-icon"><CalendarDays size={18}/></div></div><div className="stat-value">{displaySchedules.length}</div><div className="stat-label">Scheduled classes</div></div><div className="panel stat"><div className="stat-icon"><BookOpen size={18}/></div><div className="stat-value">{displaySyllabus.length}</div><div className="stat-label">Courses in syllabus</div></div><div className="panel stat"><div className="stat-icon"><Users size={18}/></div><div className="stat-value">{displayStudents.length}</div><div className="stat-label">Student records</div></div><div className="panel stat"><div className="stat-icon"><FileText size={18}/></div><div className="stat-value">{displayUploads.filter(u=>u.status==='approved').length}</div><div className="stat-label">Approved reports</div></div></div>
          <div className="section-grid"><div className="panel panel-pad"><div className="section-head"><div><div className="panel-heading">Monday's classes</div><p className="panel-sub mt-1">A look at the start of your week.</p></div><span className="day-chip">MONDAY</span></div>{displaySchedules.filter(s=>s.day_of_week===1).slice(0,3).map(s=><div className="schedule-row" key={s.id}><div className="schedule-time">{labelTime(s.start_time)}<small>{labelTime(s.end_time)}</small></div><div className="schedule-line"/><div><div className="schedule-subject">{s.subject}</div><div className="schedule-meta">{s.section_name} · {s.teacher_name}</div></div><div className="room">{s.room_name}</div></div>)}{!displaySchedules.some(s=>s.day_of_week===1)&&<Empty label="No classes scheduled" detail="Monday is currently clear."/>}</div><div className="panel panel-pad"><div className="panel-heading">Grade release</div><p className="panel-sub mt-1">From teacher submission to family access.</p><div className="notice"><ShieldCheck size={19} className="shrink-0"/><div><strong className="text-xs">Approved before release</strong><p>Excel grades are averaged automatically. Students and parents see results only after an adviser approves them.</p></div></div><div className="mt-5 flex items-center gap-2 text-xs text-muted-foreground"><FileSpreadsheet size={16} className="text-primary"/> {displayUploads.length} grade upload{displayUploads.length===1?'':'s'} on record</div><Button variant="link" className="mt-3 !px-0 text-xs" onClick={()=>go('reports')}>Open reports <ChevronRight size={14}/></Button></div></div>
        </>}
        {section==='schedule'&&<><div className="toolbar"><div className="segmented">{days.map((d,i)=><Button key={d} variant="ghost" className={`segment ${day===i+1?'selected':''}`} onClick={()=>setDay(i+1)}>{d.slice(0,3)}</Button>)}</div>{role==='admin'&&<Button onClick={()=>open('schedule')}><Plus/> Add class</Button>}</div><div className="panel panel-pad"><div className="section-head"><div><div className="panel-heading">{days[day-1]} timetable</div><p className="panel-sub mt-1">{displaySchedules.filter(s=>s.day_of_week===day).length} classes scheduled</p></div><span className="day-chip">{days[day-1].toUpperCase()}</span></div>{displaySchedules.filter(s=>s.day_of_week===day).map(s=><div className="schedule-row" key={s.id}><div className="schedule-time">{labelTime(s.start_time)}<small>{labelTime(s.end_time)}</small></div><div className="schedule-line"/><div><div className="schedule-subject">{s.subject}</div><div className="schedule-meta">{s.section_name} · {s.teacher_name}</div></div><div className="room">{s.room_name}</div></div>)}{!displaySchedules.some(s=>s.day_of_week===day)&&<Empty label="No classes scheduled" detail="There are no classes for this day."/>}</div><div className="notice max-w-2xl"><ShieldCheck size={19} className="shrink-0"/><div><strong className="text-xs">Conflict-aware scheduling</strong><p>Classes cannot be saved when a room, teacher, or section already has an overlapping time slot.</p></div></div></>}
        {section==='reports'&&<><div className="toolbar"><div className="flex items-center gap-2"><div className="panel-heading">Grade records</div><span className="badge">{visibleGrades.length} entries</span></div><div className="flex flex-wrap gap-2">{role==='admin'&&<Button variant="outline" onClick={()=>open('student')}><Plus/> Add student</Button>}{canUpload&&<Button onClick={()=>open('upload')}><UploadCloud/> Upload Excel</Button>}{(role==='student'||role==='parent')&&displayStudents.length>0&&<Button onClick={()=>open('request')}><FileText/> Request document</Button>}{visibleGrades.length>0&&<Button variant="outline" title="Download report" aria-label="Download report" onClick={downloadReport}><Download/></Button>}</div></div><div className="panel"><div className="flex flex-wrap items-center gap-3 border-b px-5 py-4"><div className="relative flex-1 min-w-40"><Search size={15} className="absolute left-3 top-2.5 text-muted-foreground"/><input className="field !pl-9" placeholder="Search student or subject" value={search} onChange={e=>setSearch(e.target.value)} aria-label="Search grades"/></div><select className="field !w-36" value={period} onChange={e=>setPeriod(e.target.value)} aria-label="Filter report status"><option>All</option><option>Approved</option><option>Pending</option></select></div><div className="table-wrap"><table className="data-table"><thead><tr><th>Student</th><th>Subject</th><th>Prelim</th><th>Midterm</th><th>Final</th><th>Average</th><th>Status</th></tr></thead><tbody>{visibleGrades.filter(g=>{const status=displayUploads.find(u=>u.id===g.upload_id)?.status??'';const student=displayStudents.find(s=>s.id===g.student_id);return (`${student?.full_name} ${student?.student_number} ${g.subject}`).toLowerCase().includes(search.toLowerCase())&&(period==='All'||status===period.toLowerCase());}).map(g=>{const student=displayStudents.find(s=>s.id===g.student_id);const status=displayUploads.find(u=>u.id===g.upload_id)?.status??'pending';return <tr key={g.id}><td><strong>{student?.full_name??'Student'}</strong><div className="text-[10px] text-muted-foreground mt-1">{student?.student_number}</div></td><td>{g.subject}</td><td>{g.prelim}</td><td>{g.midterm}</td><td>{g.final}</td><td><strong>{g.computed_grade?.toFixed(2)??'—'}</strong></td><td><span className={`badge ${status==='pending'?'pending':''}`}>{status}</span></td></tr>})}</tbody></table>{visibleGrades.length===0&&<Empty label="No grade records yet" detail={isStaff?'Upload an Excel file to get started.':'Approved grades will appear here when released.'}/>}</div></div>
          {isStaff&&<div className="panel panel-pad mt-5"><div className="section-head"><div><div className="panel-heading">Excel submissions</div><p className="panel-sub mt-1">Averages use the mean of prelim, midterm, and final scores.</p></div><FileSpreadsheet size={19} className="text-primary"/></div>{displayUploads.length===0?<Empty label="No submissions yet" detail="Teacher uploads will appear here."/>:displayUploads.map(u=><div key={u.id} className="flex flex-wrap items-center justify-between gap-3 border-t py-3"><div><div className="text-xs font-bold">{u.subject} <span className="font-normal text-muted-foreground">· {u.section_name}</span></div><div className="text-[11px] text-muted-foreground mt-1">{u.file_name} · {u.uploaded_at.slice(0,10)}</div></div><div className="flex items-center gap-2"><span className={`badge ${u.status==='pending'?'pending':''}`}>{u.status}</span>{u.status==='pending'&&(role==='adviser'||role==='admin')&&<Button size="sm" disabled={busy} onClick={()=>void approve(u)}>Approve</Button>}</div></div>)}</div>}
          {(role==='student'||role==='parent')&&requests.length>0&&<div className="panel panel-pad mt-5"><div className="panel-heading mb-3">Document requests</div>{requests.map(r=><div key={r.id} className="flex justify-between border-t py-3 text-xs"><span>{r.document_type}</span><span className="badge pending">{r.status}</span></div>)}</div>}
          <div className="notice max-w-2xl"><CircleHelp size={19} className="shrink-0"/><div><strong className="text-xs">How the average is calculated</strong><p>Prelim + Midterm + Final, divided by three. Results are shown to two decimal places. Only approved grades are visible to students and parents.</p></div></div>
        </>}
        {section==='syllabus'&&<><div className="toolbar"><div className="panel-heading">Subjects this term</div>{(role==='admin'||role==='teacher')&&<Button onClick={()=>open('syllabus')}><Plus/> Add syllabus</Button>}</div><div className="syllabus-grid">{displaySyllabus.map(s=><div className="panel syllabus-card" key={s.id}><div className="stat-icon"><BookOpen size={18}/></div><h3>{s.subject}</h3><div className="text-[11px] text-muted-foreground mt-1">{s.section_name} · {s.teacher_name}</div><p>{s.overview}</p><div className="topics"><div className="eyebrow">Course topics</div><p>{s.topics}</p></div></div>)}</div>{displaySyllabus.length===0&&<div className="panel mt-5"><Empty label="No syllabus added yet" detail="Course plans will appear here once published."/></div>}</>}
      </div>
    </main>
    <div className="mobile-nav">{nav.map(item=><Button key={item.id} variant="ghost" className={`mobile-item ${section===item.id?'active':''}`} onClick={()=>go(item.id)}><item.icon/>{item.label.split(' ')[0]}</Button>)}</div>
    {modal&&<div className="modal-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget)close()}}><div className="modal" role="dialog" aria-modal="true" aria-label={modal==='auth'?'Sign in':`Add ${modal}`}><div className="flex items-start justify-between gap-3"><div><h2>{modal==='auth'?(authMode==='sign-in'?'Welcome back':'Create an account'):modal==='upload'?'Upload grade workbook':modal==='schedule'?'Add a class':modal==='syllabus'?'Add a syllabus':modal==='student'?'Register a student':'Request a document'}</h2><p className="panel-sub mt-1">{modal==='auth'?'Access your school records securely.':modal==='upload'?'Import scores for one subject and section.':'Fill in the details below.'}</p></div><Button variant="ghost" size="icon" aria-label="Close" onClick={close}><X/></Button></div>
      {notice&&<div className="inline-alert" role="status">{notice}</div>}
      {modal==='auth'&&<><form className="mt-6 grid gap-3" onSubmit={authSubmit}>{authMode==='sign-up'&&<Field label="Full name" name="name" required/>}<Field label="Email address" name="email" type="email" required/><Field label="Password" name="password" type="password" required minLength={6}/><Button className="mt-2" disabled={busy}>{busy?'Please wait…':authMode==='sign-in'?'Sign in':'Create account'}</Button></form><div className="my-4 flex items-center gap-3 text-xs text-muted-foreground"><div className="h-px bg-border flex-1"/>or<div className="h-px bg-border flex-1"/></div><Button variant="outline" className="w-full" onClick={()=>void googleSignIn()} disabled={busy}>Continue with Google</Button><Button variant="link" className="w-full mt-3" onClick={()=>{setAuthMode(authMode==='sign-in'?'sign-up':'sign-in');setNotice('')}}>{authMode==='sign-in'?'New here? Create an account':'Already have an account? Sign in'}</Button></>}
      {modal==='upload'&&<form className="mt-6 grid gap-3" onSubmit={uploadExcel}><Field label="Subject" name="subject" placeholder="e.g. Mathematics 10" required/><Field label="Section" name="section" placeholder="e.g. Grade 10 - St. Anne" required/><label className="form-field">Excel workbook (.xlsx, .xls)<input className="field" ref={fileRef} type="file" accept=".xlsx,.xls" required/></label><p className="form-help">First row headings: <strong>Student Number, Prelim, Midterm, Final</strong>. Students must already be registered in the selected section. Grades stay pending until an adviser approves them.</p><div className="form-actions"><Button type="button" variant="outline" onClick={close}>Cancel</Button><Button disabled={busy}><UploadCloud/> {busy?'Importing…':'Import grades'}</Button></div></form>}
      {modal==='schedule'&&<form className="mt-6 grid gap-3" onSubmit={addSchedule}><div className="form-grid"><Field label="Subject" name="subject" required/><Field label="Section" name="section" required/><Field label="Teacher" name="teacher" required/><Field label="Room" name="room" required/><label className="form-field">Day<select className="field" name="day">{days.map((d,i)=><option value={i+1} key={d}>{d}</option>)}</select></label><div/></div><div className="form-grid"><Field label="Start time" name="start" type="time" required/><Field label="End time" name="end" type="time" required/></div><p className="form-help">Overlapping rooms, teachers, and sections are automatically blocked.</p><div className="form-actions"><Button type="button" variant="outline" onClick={close}>Cancel</Button><Button disabled={busy}>Save class</Button></div></form>}
      {modal==='syllabus'&&<form className="mt-6 grid gap-3" onSubmit={addSyllabus}><Field label="Subject" name="subject" required/><Field label="Section" name="section" required/><Field label="Teacher" name="teacher" required/><label className="form-field">Overview<textarea className="field min-h-20" name="overview" required/></label><label className="form-field">Topics<textarea className="field min-h-20" name="topics" required/></label><div className="form-actions"><Button type="button" variant="outline" onClick={close}>Cancel</Button><Button disabled={busy}>Save syllabus</Button></div></form>}
      {modal==='student'&&<form className="mt-6 grid gap-3" onSubmit={addStudent}><Field label="Student number" name="number" required/><Field label="Full name" name="name" required/><Field label="Section" name="section" required/><p className="form-help">Link student and parent accounts to this record in the school administration records to grant private access.</p><div className="form-actions"><Button type="button" variant="outline" onClick={close}>Cancel</Button><Button disabled={busy}>Register student</Button></div></form>}
      {modal==='request'&&<form className="mt-6 grid gap-3" onSubmit={requestDocument}><label className="form-field">Student<select className="field" name="student" required>{displayStudents.map(s=><option key={s.id} value={s.id}>{s.full_name}</option>)}</select></label><label className="form-field">Document<select className="field" name="type"><option>Report card</option><option>Transcript</option><option>Certificate</option></select></label><div className="form-actions"><Button type="button" variant="outline" onClick={close}>Cancel</Button><Button disabled={busy}>Submit request</Button></div></form>}
    </div></div>}
  </div>;
}
function Field({label,name,type='text',...props}:{label:string;name:string;type?:string;required?:boolean;placeholder?:string;minLength?:number}) {return <label className="form-field">{label}<input className="field" name={name} type={type} {...props}/></label>}
function Empty({label,detail}:{label:string;detail:string}) {return <div className="empty-state"><ClipboardList size={24}/><strong>{label}</strong><span>{detail}</span></div>}
