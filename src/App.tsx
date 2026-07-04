import { type FormEvent, useEffect, useMemo, useRef, useState } from 'react'
import {
  CalendarDays,
  Check,
  Download,
  ExternalLink,
  Gauge,
  Link,
  LogOut,
  Plus,
  RotateCcw,
  Search,
  SlidersHorizontal,
  Trash2,
  Upload,
} from 'lucide-react'
import * as XLSX from 'xlsx'

type Difficulty = 'easy' | 'medium' | 'hard' | 'very-hard'
type Course = {
  id: string
  subject: string
  pdfUrl: string
  pages: number
  pass1Hours: number | null
  seriesHours: number | null
  flashcards: boolean
  lastReview: string
  difficulty: Difficulty
  remarks: string
}

type PlanItem = {
  id: string
  date: string
  courseId: string
  hours: number
  note: string
  done: boolean
}

type Settings = {
  targetDate: string
  pagesPerHour: number
  dailyHours: number
  reviewEveryDays: number
  customPlan: PlanItem[]
}

type User = {
  id: number
  name: string
  email: string
}

type AuthMode = 'login' | 'register'
type SaveState = 'idle' | 'saving' | 'saved' | 'error'

const TOKEN_KEY = 'ecn-revision-token-v1'
const API_URL = import.meta.env.VITE_API_URL || '/api'
const LINA_EMAIL = 'lina@gmail.com'

const courseRows: Array<[string, number, number | null, string, string]> = [
  ['Upper Resp Infections', 51, null, '', ''],
  ['Nasopharyngeal Cancer', 24, null, '', ''],
  ['Asthma', 60, null, '', ''],
  ['Bronchiolitis', 13, null, '2026-06-05', 'serie: many numerical values -->FCs++'],
  ['COPD', 49, null, '', ''],
  ['Lower Resp infections', 27, null, '', ''],
  ['Cancer broncho pulmonaire', 40, null, '', ''],
  ['Pulmonary Tuberculosis', 33, 9, '', ''],
  ['Hydatid cyst (pulmonary)', 52, null, '', ''],
  ['Hydatid cyst (Hepatic)', 52, null, '', ''],
  ['Viral Hepatitis', 29, null, '', ''],
  ['Jaundice', 51, null, '', ''],
  ['hypothyroidism', 37, null, '', ''],
  ['hyperthyroidism', 21, null, '', ''],
  ['Red Eye', 16, null, '', ''],
  ['Pre-eclampsia & eclampsia', 23, null, '', ''],
  ['Metrorrhagia', 16, null, '', ''],
  ['Ectopic Pregnancy', 17, null, '', ''],
  ['Contraception', 15, null, '', ''],
  ['Acute Pain Management', 41, null, '', ''],
  ['Mood disorders', 37, null, '', ''],
  ['Anxiety disorders', 24, null, '', ''],
  ['Delirium', 12, null, '', ''],
  ['Schizophrenia', 26, null, '', ''],
  ['Cervical cancer', 28, null, '', ''],
  ['STIs', 38, null, '', ''],
  ['Arthrite Septique', 20, 4.5, '', 'ATB choice according to germ (cf ATB choice diagram+++)'],
  ['Rhumatoid Arthritis', 19, 7.5, '', ''],
  ['Vaccination', 29, null, '', ''],
  ['Endocardite infectieuse', 52, null, '', ''],
  ['Meningitis', 23, null, '', ''],
  ['Brain stroke', 55, null, '', ''],
  ['Headaches', 29, null, '', ''],
  ['Epilepsy', 22, null, '', ''],
  ['Coma', 26, null, '', ''],
  ['Head trauma', 42, null, '', ''],
  ['Intox CO, OP, PT', 47, null, '', ''],
  ['Urinary Lithiasis', 38, null, '', ''],
  ['urinary infections', 36, null, '', ''],
  ['Prostate tumors', 71, null, '', ''],
  ['Breast cancer', 22, null, '', ''],
  ['Pediatric Acute Dehydration', 22, null, '', ''],
  ['Acute Kidney Injruy', 49, 20, '', '1) wording discrepancies fr-eng: practice qs + cross_checking french, 2) make flashcards for the vaguely named concepts'],
  ['Hematuria', 20, null, '', ''],
  ['Edema', 21, null, '', ''],
  ['Dyskalemia', 24, null, '', ''],
  ['Acid-Base Disorders', 36, null, '', ''],
  ['Fluid disorders', 57, 11, '', ''],
  ['Hypercalcemia', 28, null, '', ''],
  ['Acute Adrenal Insufficiency', 32, null, '', ''],
  ['Diabetes', 77, null, '', ''],
  ['Dyslipidemia', 44, null, '2026-05-16', 'nomenclature of apoproteins,friedrich  classification '],
  ['Coronary Syndromes', 41, null, '', ''],
  ['Hypertension', 59, null, '', ''],
  ['Acute Chest Pain', 47, null, '', ''],
  ['Thrombo-Embolism', 46, null, '', ''],
  ['Cardiocirculatory Arrest', 33, null, '', ''],
  ['Cardiogenic shock', 35, null, '', ''],
  ['Severe Septic States', 36, null, '', '(in pass2, try it with ID)'],
  ['Hemorrhagic shock', 23, null, '', ''],
  ['Transfusion', 21, 8.4, '2026-06-15', ''],
  ['Polytrauma', 23, null, '', ''],
  ['Open Leg fractures', 24, null, '', ''],
  ['Acute Limb Ischemia', 26, null, '', ''],
  ['Purpuras', 26, 9.5, '', 'drew a mind map to organize etiologies by mechanism'],
  ['Splenomegaly', 16, null, '', ''],
  ['Superficial Adenopathies', 20, 6, '', ''],
  ['Anemia', 81, null, '', ''],
  ['GI Bleeding', 35, null, '', ''],
  ['Gastric Ulcer', 31, null, '', ''],
  ['Acute Peritonitis', 24, null, '', ''],
  ['Acute Appendicitis', 20, null, '', ''],
  ['Acute Intestinal Occlusion', 34, null, '', ''],
  ['Colorectal cancer', 37, null, '', ''],
  ['Dysphagia', 25, null, '', ''],
  ['Chronic Diarrhea', 39, null, '', ''],
]

const coursePdfUrls = [
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%203/42_infections%20des%20VAS%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/12-cancer_du_cavum_2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%204/7_Asthme%20de%20l%E2%80%99adulte%20et%20de%20l%E2%80%99enfant%202025bis.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/8-bronchiolite%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/09_%20BPCO%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%203/43_IRB%202025%20.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/11_CBP%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%206/72-TBC%20pulmonaire%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/36_hydatidose-h%C3%A9patiques%20et%20pulmonaires%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/36_hydatidose-h%C3%A9patiques%20et%20pulmonaires%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/35-%20H%C3%A9patites%20virales%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/41-Ict%C3%A8re%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/40%20hypothyroidie%20%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/39_%20Hyperthyroidie%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%206/56_oeil_rouge%202025%20.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/60%20PE%20%26%20eclampsie%202025%20.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%206/53_metrorragies%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/32_grossesse-extrauterine-2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/18_Contraception%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/61_Douleur%20AIGUE2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/70%20troubles%20de%20l\'humeur%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/69%20Trouble%20anxieux_2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/29_Etats%20confusionnels%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/63-Schizophr%C3%A9nie2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/13_COL%20UTERIN%20%20%20%20mai%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%203/44_IST%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/6_arthrite%20septique2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%206/58%20polyarthrite%20rhumatoide2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/75_vaccinations%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%204/25_Endocardite%20infectieuse2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%206/52-m%C3%A9ningite2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/1_AVC%202025%20.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/16_Les%20c%C3%A9phal%C3%A9es%20%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/26_EPILEPSIE%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/17-coma2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/67-Traumatisme%20Cranien%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/48%20intoxications%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%206/50_%20lithiase%20urinaire%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%203/45_Infection-urinaire%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/73_tumeurs%20prostatique%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/14-cancer-sein2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/19_%20la%20deshydratation%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%203/46_Insuffisance-r%C3%A9nale-aigue%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/33-H%C3%A9maturie%202025%20.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%206/55%20oedemes%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%203/71_2_Dyskali%C3%A9mies%202025%20.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/68_%20cours%20trb%20acide%20Base%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%203/71_1troubles%20de%20l\'hydratation%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/37_HYPERCALCEMIES%202025%20.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%203/47_ISA2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/20-Diab%C3%A9te%202025%20.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/23-dyslipid%C3%A9mies%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/65%20-%20SCA%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%205/38_HTA%20Residanat%202025%20.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/22_DOULEURS%20THORACIQUES%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%204/51-MVTE2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/5%20ACR%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/27_etat%20de%20choc%20cardiogenique%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/30_Etats%20septiques%20graves%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/28%20EDC%20hemorragique%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/66_TRANSFUSION_2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/59%20polytraumatisme%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/31%20Fracture%20jambe%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%206/49_ischemie%20MI%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/62-purpura%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/64-splenomegalie%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/2_ADP%202025%20.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/3-anemie%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/34_hemorragie%20digestive%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%203/74-Ulc%C3%A8re%20GD%20-%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%206/57%20peritonite%20aigue%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/4_Appendicite%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%206/54_occlusion_intestinale_aigue_2025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%205/15-%20Cancer%20colo%20Rectale%202025%20.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/24_Dysphagie%202025.pdf',
  'https://www.medecinesfax.org/useruploads/files/ECN%202025/Partie%202/21_dia%20chro%202025.pdf',
]

const courseNumbers: Record<string, string> = {
  'Upper Resp Infections': '42',
  'Nasopharyngeal Cancer': '12',
  Asthma: '7',
  Bronchiolitis: '8',
  COPD: '9',
  'Lower Resp infections': '43',
  'Cancer broncho pulmonaire': '11',
  'Pulmonary Tuberculosis': '72',
  'Hydatid cyst (pulmonary)': '36',
  'Hydatid cyst (Hepatic)': '36',
  'Viral Hepatitis': '35',
  Jaundice: '41',
  hypothyroidism: '40',
  hyperthyroidism: '39',
  'Red Eye': '56',
  'Pre-eclampsia & eclampsia': '60',
  Metrorrhagia: '53',
  'Ectopic Pregnancy': '32',
  Contraception: '19',
  'Acute Pain Management': '61',
  'Mood disorders': '70',
  'Anxiety disorders': '69',
  Delirium: '29',
  Schizophrenia: '63',
  'Cervical cancer': '13',
  STIs: '44',
  'Arthrite Septique': '6',
  'Rhumatoid Arthritis': '58',
  Vaccination: '75',
  'Endocardite infectieuse': '25',
  Meningitis: '52',
  'Brain stroke': '1',
  Headaches: '16',
  Epilepsy: '26',
  Coma: '17',
  'Head trauma': '67',
  'Intox CO, OP, PT': '48',
  'Urinary Lithiasis': '50',
  'urinary infections': '45',
  'Prostate tumors': '73',
  'Breast cancer': '14',
  'Acute Kidney Injury': '46',
  'Acute Kidney Injruy': '46',
  'Pediatric Acute Dehydration': '18',
  Hematuria: '33',
  Edema: '55',
  Dyskalemia: '71-2',
  'Acid-Base Disorders': '68',
  'Fluid disorders': '71-1',
  Hypercalcemia: '37',
  'Acute Adrenal Insufficiency': '47',
  Diabetes: '20',
  Dyslipidemia: '23',
  'Coronary Syndromes': '65',
  Hypertension: '38',
  'Acute Chest Pain': '22',
  'Thrombo-Embolism': '51',
  'Cardiocirculatory Arrest': '5',
  'Cardiogenic shock': '27',
  'Severe Septic States': '30',
  'Hemorrhagic shock': '28',
  Transfusion: '66',
  Polytrauma: '59',
  'Open Leg fractures': '31',
  'Acute Limb Ischemia': '49',
  Purpuras: '62',
  Splenomegaly: '64',
  'Superficial Adenopathies': '2',
  Anemia: '3',
  'GI Bleeding': '34',
  'Gastric Ulcer': '74',
  'Acute Peritonitis': '57',
  'Acute Appendicitis': '4',
  'Acute Intestinal Occlusion': '54',
  'Colorectal cancer': '15',
  Dysphagia: '24',
  'Chronic Diarrhea': '21',
}

function normalizeSubjectName(subject: string) {
  return subject
    .replace(/^\s*\d+(?:-\d+)?\s*-\s*/, '')
    .replace(/\s+/g, ' ')
    .replace(/\u00a0/g, ' ')
    .trim()
    .replace('Acute Kidney Injruy', 'Acute Kidney Injury')
}

function normalizeUrl(url: string) {
  const trimmed = url.trim()
  if (!trimmed) return ''
  if (/^(https?:\/\/|mailto:|tel:)/i.test(trimmed)) return trimmed
  return `https://${trimmed}`
}

function subjectWithNumber(subject: string) {
  const clean = normalizeSubjectName(subject)
  const number = courseNumbers[clean]
  return number ? `${number} - ${clean}` : subject.trim()
}

function courseWithNumber(course: Course): Course {
  const subject = subjectWithNumber(course.subject)
  return { ...course, subject, pdfUrl: course.pdfUrl || '' }
}

const initialCourses: Course[] = courseRows.map(([subject, pages, pass1Hours, lastReview, remarks], index) => ({
  id: `course-${index + 1}`,
  subject: subjectWithNumber(subject),
  pdfUrl: coursePdfUrls[index] || '',
  pages,
  pass1Hours,
  seriesHours: null,
  flashcards: false,
  lastReview,
  difficulty: 'medium',
  remarks,
}))

const defaultSettings: Settings = {
  targetDate: '2026-09-15',
  pagesPerHour: 5.5,
  dailyHours: 7,
  reviewEveryDays: 14,
  customPlan: [],
}

function defaultCoursesForUser(user: User | null) {
  return user?.email.toLowerCase() === LINA_EMAIL ? initialCourses : []
}

function isSeedCourseSet(courses: Course[]) {
  if (courses.length !== initialCourses.length) return false
  const seedSubjects = new Set(initialCourses.map((course) => normalizeSubjectName(course.subject)))
  return courses.every((course) => seedSubjects.has(normalizeSubjectName(course.subject)))
}

function round(value: number, precision = 1) {
  const factor = 10 ** precision
  return Math.round(value * factor) / factor
}

function todayIso() {
  return new Date().toISOString().slice(0, 10)
}

function daysBetween(fromIso: string, toIso: string) {
  const from = new Date(`${fromIso}T00:00:00`)
  const to = new Date(`${toIso}T00:00:00`)
  return Math.max(1, Math.ceil((to.getTime() - from.getTime()) / 86400000) + 1)
}

function addDays(iso: string, days: number) {
  const date = new Date(`${iso}T00:00:00`)
  date.setDate(date.getDate() + days)
  return date.toISOString().slice(0, 10)
}

function weekStartIso(iso: string) {
  const date = new Date(`${iso}T00:00:00`)
  const day = date.getDay() || 7
  date.setDate(date.getDate() - day + 1)
  return date.toISOString().slice(0, 10)
}

function shortDateLabel(iso: string) {
  return new Intl.DateTimeFormat('fr-FR', { weekday: 'short', day: '2-digit', month: '2-digit' }).format(new Date(`${iso}T00:00:00`))
}

function plannedHours(course: Course, settings: Settings) {
  return round(course.pages / Math.max(settings.pagesPerHour, 0.1), 1)
}

function totalCourseHours(course: Course, settings: Settings) {
  return course.pass1Hours === null ? plannedHours(course, settings) : 0
}

function statusFor(course: Course, settings: Settings) {
  if (course.pass1Hours === null) return 'pending'
  const delta = round(course.pass1Hours - plannedHours(course, settings), 1)
  if (delta > 1) return 'late'
  if (delta >= 0) return 'watch'
  return 'fast'
}

function parseNumber(value: unknown) {
  if (value === '' || value === null || value === undefined) return null
  const parsed = Number(String(value).replace(',', '.'))
  return Number.isFinite(parsed) ? parsed : null
}

function normalizeDifficulty(value: unknown): Difficulty {
  const raw = String(value || '').toLowerCase().trim()
  if (['easy', 'medium', 'hard', 'very-hard'].includes(raw)) return raw as Difficulty
  if (raw.includes('very') || raw.includes('tres') || raw.includes('très')) return 'very-hard'
  if (raw.includes('hard') || raw.includes('difficile')) return 'hard'
  if (raw.includes('easy') || raw.includes('facile')) return 'easy'
  return 'medium'
}

async function apiRequest<T>(path: string, options: RequestInit = {}, token?: string) {
  const response = await fetch(`${API_URL}${path}`, {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      ...(token ? { Authorization: `Bearer ${token}` } : {}),
      ...options.headers,
    },
  })
  const data = (await response.json().catch(() => ({}))) as T & { error?: string }
  if (!response.ok) throw new Error(data.error || 'Erreur serveur.')
  return data
}

function App() {
  const fileInputRef = useRef<HTMLInputElement>(null)
  const didLoadProgress = useRef(false)
  const [settings, setSettings] = useState<Settings>(defaultSettings)
  const [courses, setCourses] = useState<Course[]>([])
  const [query, setQuery] = useState('')
  const [filter, setFilter] = useState<'all' | 'todo' | 'done' | 'review' | 'over'>('all')
  const [isAddingPlanItem, setIsAddingPlanItem] = useState(false)
  const [planDraft, setPlanDraft] = useState<Omit<PlanItem, 'id'>>({
    date: todayIso(),
    courseId: '',
    hours: 1,
    note: '',
    done: false,
  })
  const [token, setToken] = useState(() => localStorage.getItem(TOKEN_KEY) || '')
  const [user, setUser] = useState<User | null>(null)
  const [authMode, setAuthMode] = useState<AuthMode>('login')
  const [authForm, setAuthForm] = useState({ name: '', email: '', password: '' })
  const [authError, setAuthError] = useState('')
  const [isAuthLoading, setIsAuthLoading] = useState(Boolean(token))
  const [saveState, setSaveState] = useState<SaveState>('idle')

  useEffect(() => {
    if (!token) {
      setUser(null)
      setIsAuthLoading(false)
      didLoadProgress.current = false
      return
    }

    let cancelled = false
    async function loadSession() {
      setIsAuthLoading(true)
      try {
        const session = await apiRequest<{ user: User }>('/auth/me', {}, token)
        const progress = await apiRequest<{ settings: Settings | null; courses: Course[] | null }>('/progress', {}, token)
        if (cancelled) return
        setUser(session.user)
        setSettings(progress.settings ? { ...defaultSettings, ...progress.settings } : defaultSettings)
        const loadedCourses = Array.isArray(progress.courses)
          ? progress.courses.map(courseWithNumber)
          : defaultCoursesForUser(session.user)
        const shouldClearSeedForOtherUser =
          session.user.email.toLowerCase() !== LINA_EMAIL && isSeedCourseSet(loadedCourses)
        setCourses(shouldClearSeedForOtherUser ? [] : loadedCourses)
        didLoadProgress.current = true
      } catch (error) {
        if (cancelled) return
        localStorage.removeItem(TOKEN_KEY)
        setToken('')
        setUser(null)
        setAuthError(error instanceof Error ? error.message : 'Session invalide.')
      } finally {
        if (!cancelled) setIsAuthLoading(false)
      }
    }
    void loadSession()

    return () => {
      cancelled = true
    }
  }, [token])

  useEffect(() => {
    if (!token || !user || !didLoadProgress.current) return
    setSaveState('saving')
    const timeoutId = window.setTimeout(() => {
      apiRequest('/progress', {
        method: 'PUT',
        body: JSON.stringify({ settings, courses }),
      }, token)
        .then(() => setSaveState('saved'))
        .catch(() => setSaveState('error'))
    }, 650)

    return () => window.clearTimeout(timeoutId)
  }, [courses, settings, token, user])

  const stats = useMemo(() => {
    const totalPages = courses.reduce((sum, course) => sum + course.pages, 0)
    const completedCourses = courses.filter((course) => course.pass1Hours !== null)
    const done = completedCourses.length
    const passPlanned = courses.reduce((sum, course) => sum + plannedHours(course, settings), 0)
    const passPlannedDone = completedCourses.reduce((sum, course) => sum + plannedHours(course, settings), 0)
    const passDone = completedCourses.reduce((sum, course) => sum + (course.pass1Hours || 0), 0)
    const passDelta = passDone - passPlannedDone
    const remaining = courses.reduce((sum, course) => sum + totalCourseHours(course, settings), 0)
    const customPlanned = settings.customPlan.reduce((sum, item) => sum + item.hours, 0)
    const customDone = settings.customPlan.reduce((sum, item) => sum + (item.done ? item.hours : 0), 0)
    const daysToTarget = daysBetween(todayIso(), settings.targetDate)
    const dailyRequired = remaining / daysToTarget
    const finishDays = Math.ceil(remaining / Math.max(settings.dailyHours, 0.1))
    const predictedFinish = addDays(todayIso(), Math.max(0, finishDays - 1))
    return {
      totalPages,
      done,
      passPlanned: round(passPlanned, 1),
      passPlannedDone: round(passPlannedDone, 1),
      passDone: round(passDone, 1),
      passDelta: round(passDelta, 1),
      passAverageDelta: done ? round(passDelta / done, 1) : 0,
      remaining: round(remaining, 1),
      customPlanned: round(customPlanned, 1),
      customDone: round(customDone, 1),
      customGap: round(customPlanned - remaining, 1),
      progress: courses.length ? Math.round((done / courses.length) * 100) : 0,
      daysToTarget,
      dailyRequired: round(dailyRequired, 1),
      predictedFinish,
      isOnTrack: dailyRequired <= settings.dailyHours,
    }
  }, [courses, settings])

  const filteredCourses = useMemo(() => {
    const needle = query.trim().toLowerCase()
    return courses.filter((course) => {
      const goal = plannedHours(course, settings)
      const matchesSearch = !needle || `${course.subject} ${course.remarks}`.toLowerCase().includes(needle)
      const needsReview =
        course.lastReview !== '' &&
        daysBetween(course.lastReview, todayIso()) > settings.reviewEveryDays
      const matchesFilter =
        filter === 'all' ||
        (filter === 'todo' && course.pass1Hours === null) ||
        (filter === 'done' && course.pass1Hours !== null) ||
        (filter === 'review' && needsReview) ||
        (filter === 'over' && course.pass1Hours !== null && course.pass1Hours - goal > 1)
      return matchesSearch && matchesFilter
    })
  }, [courses, filter, query, settings])

  const customSchedule = useMemo(() => (
    settings.customPlan
      .map((item) => ({ ...item, done: Boolean(item.done) }))
      .sort((a, b) => a.date.localeCompare(b.date))
  ), [settings.customPlan])

  const customWeeks = useMemo(() => {
    const weeks = new Map<string, Map<string, PlanItem[]>>()
    for (const item of customSchedule) {
      const weekKey = weekStartIso(item.date)
      if (!weeks.has(weekKey)) weeks.set(weekKey, new Map())
      const days = weeks.get(weekKey)!
      if (!days.has(item.date)) days.set(item.date, [])
      days.get(item.date)!.push(item)
    }

    return [...weeks.entries()].map(([weekStart, days]) => ({
      weekStart,
      weekEnd: addDays(weekStart, 6),
      days: [...days.entries()].map(([date, items]) => ({
        date,
        items,
        totalHours: round(items.reduce((sum, item) => sum + item.hours, 0), 1),
        doneHours: round(items.reduce((sum, item) => sum + (item.done ? item.hours : 0), 0), 1),
      })),
    }))
  }, [customSchedule])

  const unfinishedCourses = useMemo(() => courses.filter((course) => course.pass1Hours === null), [courses])

  const planCourseOptions = (selectedCourseId: string) => {
    const hasSelectedCourse = unfinishedCourses.some((course) => course.id === selectedCourseId)
    const selectedCourse = courses.find((course) => course.id === selectedCourseId)
    return hasSelectedCourse || !selectedCourse ? unfinishedCourses : [selectedCourse, ...unfinishedCourses]
  }

  const updateCourse = (id: string, patch: Partial<Course>) => {
    setCourses((current) => current.map((course) => (course.id === id ? { ...course, ...patch } : course)))
  }

  const updateSettings = (patch: Partial<Settings>) => {
    setSettings((current) => ({ ...current, ...patch }))
  }

  const editPdfUrl = (course: Course) => {
    const value = window.prompt('Lien PDF du cours', course.pdfUrl)
    if (value === null) return
    updateCourse(course.id, { pdfUrl: normalizeUrl(value) })
  }

  const startPlanItem = () => {
    setPlanDraft({
      date: todayIso(),
      courseId: unfinishedCourses[0]?.id || '',
      hours: 1,
      note: '',
      done: false,
    })
    setIsAddingPlanItem(true)
  }

  const validatePlanItem = () => {
    if (!planDraft.courseId) return
    updateSettings({
      customPlan: [
        ...settings.customPlan,
        {
          id: `plan-${Date.now()}`,
          ...planDraft,
          hours: Number(planDraft.hours) || 0,
        },
      ],
    })
    setIsAddingPlanItem(false)
  }

  const updatePlanItem = (id: string, patch: Partial<PlanItem>) => {
    updateSettings({
      customPlan: settings.customPlan.map((item) => (item.id === id ? { ...item, ...patch } : item)),
    })
  }

  const removePlanItem = (id: string) => {
    updateSettings({
      customPlan: settings.customPlan.filter((item) => item.id !== id),
    })
  }

  const addCourse = () => {
    setCourses((current) => [
      ...current,
      {
        id: `course-${Date.now()}`,
        subject: 'Nouveau cours',
        pdfUrl: '',
        pages: 0,
        pass1Hours: null,
        seriesHours: null,
        flashcards: false,
        lastReview: '',
        difficulty: 'medium',
        remarks: '',
      },
    ])
  }

  const exportExcel = () => {
    const rows = courses.map((course) => ({
      Subject: course.subject,
      'PDF URL': course.pdfUrl,
      'Page Count': course.pages,
      'Pass1 goal': plannedHours(course, settings),
      Pass1: course.pass1Hours ?? '',
      'Delta hours': course.pass1Hours === null ? '' : round(course.pass1Hours - plannedHours(course, settings), 1),
      'Last Review': course.lastReview,
      Serie: course.seriesHours ?? '',
      Flashcards: course.flashcards ? 'yes' : '',
      difficulty: course.difficulty,
      remarks: course.remarks,
    }))
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), 'ECN suivi')
    XLSX.writeFile(workbook, `ecn-revisions-${todayIso()}.xlsx`)
  }

  const importExcel = async (file: File) => {
    const buffer = await file.arrayBuffer()
    const workbook = XLSX.read(buffer, { cellDates: true })
    const sheet = workbook.Sheets[workbook.SheetNames[0]]
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' })
    const imported = rows
      .filter((row) => row.Subject || row.subject || row.Cours)
      .map((row, index): Course => {
        const lastReview = row['Last Review']
        return {
          id: `import-${Date.now()}-${index}`,
          subject: subjectWithNumber(String(row.Subject || row.subject || row.Cours || '')),
          pdfUrl: normalizeUrl(String(row['PDF URL'] || row.pdfUrl || row.PDF || '')),
          pages: parseNumber(row['Page Count'] || row.pages || row.Pages) || 0,
          pass1Hours: parseNumber(row.Pass1 || row.pass1),
          seriesHours: parseNumber(row.Serie || row.series),
          flashcards: Boolean(row.Flashcards || row.flashcards),
          lastReview:
            lastReview instanceof Date
              ? lastReview.toISOString().slice(0, 10)
              : String(lastReview || ''),
          difficulty: normalizeDifficulty(row.difficulty || row.Difficulty),
          remarks: String(row.remarks || row.Remarks || ''),
        }
      })
    if (imported.length) setCourses(imported)
  }

  const resetData = () => {
    setCourses(defaultCoursesForUser(user))
    setSettings(defaultSettings)
  }

  const submitAuth = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault()
    setAuthError('')
    setIsAuthLoading(true)
    try {
      const payload =
        authMode === 'register'
          ? authForm
          : { email: authForm.email, password: authForm.password }
      const result = await apiRequest<{ token: string; user: User }>(`/auth/${authMode}`, {
        method: 'POST',
        body: JSON.stringify(payload),
      })
      localStorage.setItem(TOKEN_KEY, result.token)
      setToken(result.token)
      setUser(result.user)
      didLoadProgress.current = false
    } catch (error) {
      setAuthError(error instanceof Error ? error.message : 'Connexion impossible.')
    } finally {
      setIsAuthLoading(false)
    }
  }

  const logout = () => {
    localStorage.removeItem(TOKEN_KEY)
    setToken('')
    setUser(null)
    setSettings(defaultSettings)
    setCourses([])
    setSaveState('idle')
  }

  if (isAuthLoading && !user) {
    return (
      <main className="authShell">
        <div className="authPanel">
          <p className="eyebrow">ECN revision cockpit</p>
          <h1>Chargement...</h1>
        </div>
      </main>
    )
  }

  if (!user) {
    return (
      <main className="authShell">
        <form className="authPanel" onSubmit={submitAuth}>
          <p className="eyebrow">ECN revision cockpit</p>
          <h1>{authMode === 'login' ? 'Connexion' : 'Creer un compte'}</h1>
          <p className="authHint">Chaque utilisateur garde son planning, ses heures, ses reviews et ses remarques dans la base SQLite.</p>

          {authMode === 'register' && (
            <label>
              Nom
              <input value={authForm.name} onChange={(event) => setAuthForm((current) => ({ ...current, name: event.target.value }))} />
            </label>
          )}
          <label>
            Email
            <input type="email" value={authForm.email} onChange={(event) => setAuthForm((current) => ({ ...current, email: event.target.value }))} />
          </label>
          <label>
            Mot de passe
            <input type="password" minLength={6} value={authForm.password} onChange={(event) => setAuthForm((current) => ({ ...current, password: event.target.value }))} />
          </label>
          {authError && <p className="authError">{authError}</p>}
          <button className="authSubmit" type="submit" disabled={isAuthLoading}>
            {isAuthLoading ? 'Patiente...' : authMode === 'login' ? 'Se connecter' : "S'inscrire"}
          </button>
          <button
            className="authSwitch"
            type="button"
            onClick={() => {
              setAuthMode((current) => (current === 'login' ? 'register' : 'login'))
              setAuthError('')
            }}
          >
            {authMode === 'login' ? 'Creer un nouveau compte' : 'Jai deja un compte'}
          </button>
        </form>
      </main>
    )
  }

  return (
    <main className="appShell">
      <header className="topBar">
        <div>
          <p className="eyebrow">ECN revision cockpit</p>
          <h1>Planifier, suivre, ajuster.</h1>
          <p className={`syncState ${saveState}`}>{saveState === 'saving' ? 'Sauvegarde...' : saveState === 'saved' ? 'Sauvegarde BDD OK' : saveState === 'error' ? 'Erreur sauvegarde' : 'Connecte'}</p>
        </div>
        <div className="topActions">
          <div className="userBadge">
            <span>{user.name}</span>
            <small>{user.email}</small>
          </div>
          <button className="iconButton" type="button" title="Importer Excel" onClick={() => fileInputRef.current?.click()}>
            <Upload size={18} />
          </button>
          <button className="iconButton" type="button" title="Exporter Excel" onClick={exportExcel}>
            <Download size={18} />
          </button>
          <button className="iconButton" type="button" title="Reinitialiser" onClick={resetData}>
            <RotateCcw size={18} />
          </button>
          <button className="iconButton" type="button" title="Deconnexion" onClick={logout}>
            <LogOut size={18} />
          </button>
          <input
            ref={fileInputRef}
            className="hiddenInput"
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={(event) => {
              const file = event.target.files?.[0]
              if (file) void importExcel(file)
              event.target.value = ''
            }}
          />
        </div>
      </header>

      <section className="settingsGrid" aria-label="Parametres">
        <label>
          <span><CalendarDays size={16} /> Date cible</span>
          <input type="date" value={settings.targetDate} onChange={(event) => updateSettings({ targetDate: event.target.value })} />
        </label>
        <label>
          <span><Gauge size={16} /> Pages par heure</span>
          <input type="number" min="0.5" step="0.1" value={settings.pagesPerHour} onChange={(event) => updateSettings({ pagesPerHour: Number(event.target.value) })} />
        </label>
        <label>
          <span><SlidersHorizontal size={16} /> Heures par jour</span>
          <input type="number" min="0.5" step="0.5" value={settings.dailyHours} onChange={(event) => updateSettings({ dailyHours: Number(event.target.value) })} />
        </label>
      </section>

      <section className="metricsGrid" aria-label="Synthese">
        <Metric label="Progression" value={`${stats.progress}%`} hint={`${stats.done}/${courses.length} cours pass1 finis`} tone="blue" />
        <Metric
          label="Ecart Pass1"
          value={`${stats.passDelta > 0 ? '+' : ''}${stats.passDelta}h`}
          hint={`${stats.passDone}h faites / ${stats.passPlannedDone}h prevues, moy ${stats.passAverageDelta > 0 ? '+' : ''}${stats.passAverageDelta}h`}
          tone={stats.done === 0 ? 'blue' : stats.passDelta > 0 ? 'red' : 'green'}
        />
        <Metric label="Restant estime" value={`${stats.remaining}h`} hint={`${stats.totalPages} pages au total`} tone="amber" />
        <Metric label="Rythme requis" value={`${stats.dailyRequired}h/j`} hint={`jusqu'au ${settings.targetDate}`} tone={stats.isOnTrack ? 'green' : 'red'} />
        <Metric label="Fin prevue" value={stats.predictedFinish} hint={stats.isOnTrack ? 'tu es dans le rythme' : 'augmente le rythme ou decale la cible'} tone={stats.isOnTrack ? 'green' : 'red'} />
      </section>

      <section className="workArea">
        <aside className="sidePanel">
          <div className="panelHeader">
            <h2>Planning</h2>
            <span>{stats.customPlanned}h</span>
          </div>

          <div className="customPlanner">
            <div className="customPlanSummary">
              <strong>{stats.customPlanned}h planifiees</strong>
              <span>{stats.customDone}h cochees | {stats.customGap >= 0 ? '+' : ''}{stats.customGap}h vs restant</span>
            </div>

            {!isAddingPlanItem ? (
              <button className="primaryButton fullWidth" type="button" onClick={startPlanItem}>
                <Plus size={17} /> Ajouter
              </button>
            ) : (
              <div className="planItem draftPlanItem">
                <div className="planItemTop">
                  <input type="date" value={planDraft.date} onChange={(event) => setPlanDraft((current) => ({ ...current, date: event.target.value }))} />
                  <input type="number" min="0" step="0.25" value={planDraft.hours} onChange={(event) => setPlanDraft((current) => ({ ...current, hours: Number(event.target.value) }))} />
                  <button className="iconButton compact" type="button" title="Annuler" onClick={() => setIsAddingPlanItem(false)}>
                    <Trash2 size={15} />
                  </button>
                </div>
                <select value={planDraft.courseId} onChange={(event) => setPlanDraft((current) => ({ ...current, courseId: event.target.value }))}>
                  <option value="">Choisir un cours</option>
                  {planCourseOptions(planDraft.courseId).map((course) => (
                    <option key={course.id} value={course.id}>{course.subject}</option>
                  ))}
                </select>
                <input placeholder="Note" value={planDraft.note} onChange={(event) => setPlanDraft((current) => ({ ...current, note: event.target.value }))} />
                <button className="primaryButton fullWidth" type="button" disabled={!planDraft.courseId} onClick={validatePlanItem}>
                  <Check size={17} /> Valider
                </button>
              </div>
            )}

            <div className="planWeekList">
              {customWeeks.map((week) => (
                <section className="planWeek" key={week.weekStart}>
                  <h3>Semaine {shortDateLabel(week.weekStart)} - {shortDateLabel(week.weekEnd)}</h3>
                  {week.days.map((day) => (
                    <div className="planDay" key={day.date}>
                      <div className="planDayHeader">
                        <strong>{shortDateLabel(day.date)}</strong>
                        <span className={day.totalHours > settings.dailyHours ? 'over' : ''}>
                          {day.totalHours}/{settings.dailyHours}h
                        </span>
                      </div>
                      <div className="planItemList">
                        {day.items.map((item) => (
                          <div className={`planItem ${item.done ? 'done' : ''}`} key={item.id}>
                            <div className="planItemTop">
                              <label className="planDoneToggle">
                                <input type="checkbox" checked={item.done} onChange={(event) => updatePlanItem(item.id, { done: event.target.checked })} />
                                <span>{item.done ? 'Fait' : 'A faire'}</span>
                              </label>
                              <input type="date" value={item.date} onChange={(event) => updatePlanItem(item.id, { date: event.target.value })} />
                              <input type="number" min="0" step="0.25" value={item.hours} onChange={(event) => updatePlanItem(item.id, { hours: Number(event.target.value) })} />
                              <button className="iconButton compact" type="button" title="Supprimer la seance" onClick={() => removePlanItem(item.id)}>
                                <Trash2 size={15} />
                              </button>
                            </div>
                            <select value={item.courseId} onChange={(event) => updatePlanItem(item.id, { courseId: event.target.value })}>
                              <option value="">Choisir un cours</option>
                              {planCourseOptions(item.courseId).map((course) => (
                                <option key={course.id} value={course.id}>{course.subject}</option>
                              ))}
                            </select>
                            <input placeholder="Note" value={item.note} onChange={(event) => updatePlanItem(item.id, { note: event.target.value })} />
                            {item.note.trim() && <p>{item.note}</p>}
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </section>
              ))}
              {!customSchedule.length && <p className="empty">Ajoute tes seances avec les cours et les heures que tu veux.</p>}
            </div>
          </div>
        </aside>

        <section className="tableSection">
          <div className="tableToolbar">
            <div className="searchBox">
              <Search size={17} />
              <input placeholder="Chercher un cours ou une remarque" value={query} onChange={(event) => setQuery(event.target.value)} />
            </div>
            <select value={filter} onChange={(event) => setFilter(event.target.value as typeof filter)}>
              <option value="all">Tous</option>
              <option value="todo">A faire</option>
              <option value="done">Pass1 fini</option>
              <option value="review">Review due</option>
              <option value="over">Depassement</option>
            </select>
            <button className="primaryButton" type="button" onClick={addCourse}>
              <Plus size={17} /> Cours
            </button>
          </div>

          <div className="tableWrap">
            <table>
              <thead>
                <tr>
                  <th>Cours</th>
                  <th>Pages</th>
                  <th>Goal</th>
                  <th>Pass1</th>
                  <th>Etat</th>
                  <th>Review</th>
                  <th>Series</th>
                  <th>FC</th>
                  <th>Difficulte</th>
                  <th>Remarques</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {filteredCourses.map((course) => {
                  const goal = plannedHours(course, settings)
                  const delta = course.pass1Hours === null ? null : round(course.pass1Hours - goal, 1)
                  const status = statusFor(course, settings)
                  return (
                    <tr className={course.pass1Hours !== null ? 'completedRow' : ''} key={course.id}>
                      <td className="subjectCell">
                        <div className="courseTitleRow">
                          {course.pdfUrl ? (
                            <a className="courseLink" href={course.pdfUrl} target="_blank" rel="noreferrer">
                              {course.subject}
                              <ExternalLink size={13} />
                            </a>
                          ) : (
                            <span className="courseLink disabled">{course.subject}</span>
                          )}
                          <button
                            className={`linkEditButton ${course.pdfUrl ? 'active' : ''}`}
                            type="button"
                            title={course.pdfUrl ? 'Modifier le lien PDF' : 'Ajouter un lien PDF'}
                            onClick={() => editPdfUrl(course)}
                          >
                            <Link size={14} />
                          </button>
                        </div>
                        <input
                          className="subjectEdit"
                          value={course.subject}
                          onChange={(event) => {
                            const subject = event.target.value
                            updateCourse(course.id, { subject })
                          }}
                        />
                      </td>
                      <td><NumberInput value={course.pages} onChange={(pages) => updateCourse(course.id, { pages })} /></td>
                      <td className="goalCell">{goal}h</td>
                      <td><NullableNumberInput value={course.pass1Hours} onChange={(pass1Hours) => updateCourse(course.id, { pass1Hours })} /></td>
                      <td><span className={`statusPill ${status}`}>{delta === null ? 'pending' : `${delta > 0 ? '+' : ''}${delta}h`}</span></td>
                      <td><input type="date" value={course.lastReview} onChange={(event) => updateCourse(course.id, { lastReview: event.target.value })} /></td>
                      <td><NullableNumberInput value={course.seriesHours} onChange={(seriesHours) => updateCourse(course.id, { seriesHours })} /></td>
                      <td className="checkCell">
                        <button
                          className={`checkButton ${course.flashcards ? 'active' : ''}`}
                          type="button"
                          title="Flashcards"
                          onClick={() => updateCourse(course.id, { flashcards: !course.flashcards })}
                        >
                          <Check size={15} />
                        </button>
                      </td>
                      <td>
                        <select value={course.difficulty} onChange={(event) => updateCourse(course.id, { difficulty: event.target.value as Difficulty })}>
                          <option value="easy">easy</option>
                          <option value="medium">medium</option>
                          <option value="hard">hard</option>
                          <option value="very-hard">very hard</option>
                        </select>
                      </td>
                      <td className="remarksCell">
                        <textarea value={course.remarks} onChange={(event) => updateCourse(course.id, { remarks: event.target.value })} />
                      </td>
                      <td>
                        <button className="iconButton compact" type="button" title="Supprimer" onClick={() => setCourses((current) => current.filter((item) => item.id !== course.id))}>
                          <Trash2 size={15} />
                        </button>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        </section>
      </section>
    </main>
  )
}

function Metric({ label, value, hint, tone }: { label: string; value: string; hint: string; tone: 'blue' | 'amber' | 'green' | 'red' }) {
  return (
    <div className={`metric ${tone}`}>
      <span>{label}</span>
      <strong>{value}</strong>
      <p>{hint}</p>
    </div>
  )
}

function NumberInput({ value, onChange }: { value: number; onChange: (value: number) => void }) {
  return (
    <input
      className="numberInput"
      type="number"
      min="0"
      step="1"
      value={value}
      onChange={(event) => onChange(Number(event.target.value))}
    />
  )
}

function NullableNumberInput({ value, onChange }: { value: number | null; onChange: (value: number | null) => void }) {
  return (
    <input
      className="numberInput"
      type="number"
      min="0"
      step="0.1"
      placeholder="-"
      value={value ?? ''}
      onChange={(event) => onChange(parseNumber(event.target.value))}
    />
  )
}

export default App
