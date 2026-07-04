const path = require('node:path')
const express = require('express')
const cors = require('cors')
const Database = require('better-sqlite3')
const { Pool } = require('pg')
const bcrypt = require('bcryptjs')
const jwt = require('jsonwebtoken')

const PORT = Number(process.env.PORT || 3001)
const JWT_SECRET = process.env.JWT_SECRET || 'change-this-secret-before-production'
const DB_PATH = process.env.DB_PATH || path.join(__dirname, 'ecn-revisions.db')
const DATABASE_URL = process.env.DATABASE_URL

const app = express()
app.use(cors({ origin: ['http://127.0.0.1:5173', 'http://localhost:5173'] }))
app.use(express.json({ limit: '5mb' }))

function createStore() {
  if (DATABASE_URL) {
    const pool = new Pool({
      connectionString: DATABASE_URL,
      ssl: process.env.PGSSL === 'disable' ? false : { rejectUnauthorized: false },
    })

    return {
      kind: 'postgres',
      async init() {
        await pool.query(`
          CREATE TABLE IF NOT EXISTS users (
            id BIGSERIAL PRIMARY KEY,
            name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            created_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
          );

          CREATE TABLE IF NOT EXISTS progress (
            user_id BIGINT PRIMARY KEY REFERENCES users(id) ON DELETE CASCADE,
            settings_json JSONB NOT NULL,
            courses_json JSONB NOT NULL,
            updated_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
          );
        `)
      },
      async getUserById(id) {
        const result = await pool.query('SELECT id, name, email FROM users WHERE id = $1', [id])
        return result.rows[0] || null
      },
      async getUserByEmail(email) {
        const result = await pool.query('SELECT * FROM users WHERE email = $1', [email])
        return result.rows[0] || null
      },
      async createUser(name, email, passwordHash) {
        const result = await pool.query(
          'INSERT INTO users (name, email, password_hash) VALUES ($1, $2, $3) RETURNING id, name, email',
          [name, email, passwordHash],
        )
        return result.rows[0]
      },
      async getProgress(userId) {
        const result = await pool.query(
          'SELECT settings_json, courses_json, updated_at FROM progress WHERE user_id = $1',
          [userId],
        )
        return result.rows[0] || null
      },
      async saveProgress(userId, settings, courses) {
        await pool.query(
          `
            INSERT INTO progress (user_id, settings_json, courses_json, updated_at)
            VALUES ($1, $2::jsonb, $3::jsonb, CURRENT_TIMESTAMP)
            ON CONFLICT(user_id) DO UPDATE SET
              settings_json = excluded.settings_json,
              courses_json = excluded.courses_json,
              updated_at = CURRENT_TIMESTAMP
          `,
          [userId, JSON.stringify(settings), JSON.stringify(courses)],
        )
      },
    }
  }

  const db = new Database(DB_PATH)
  db.pragma('journal_mode = WAL')
  db.pragma('foreign_keys = ON')

  return {
    kind: 'sqlite',
    async init() {
      db.exec(`
        CREATE TABLE IF NOT EXISTS users (
          id INTEGER PRIMARY KEY AUTOINCREMENT,
          name TEXT NOT NULL,
          email TEXT NOT NULL UNIQUE,
          password_hash TEXT NOT NULL,
          created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS progress (
          user_id INTEGER PRIMARY KEY,
          settings_json TEXT NOT NULL,
          courses_json TEXT NOT NULL,
          updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
          FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE CASCADE
        );
      `)
    },
    async getUserById(id) {
      return db.prepare('SELECT id, name, email FROM users WHERE id = ?').get(id) || null
    },
    async getUserByEmail(email) {
      return db.prepare('SELECT * FROM users WHERE email = ?').get(email) || null
    },
    async createUser(name, email, passwordHash) {
      const result = db
        .prepare('INSERT INTO users (name, email, password_hash) VALUES (?, ?, ?)')
        .run(name, email, passwordHash)
      return db.prepare('SELECT id, name, email FROM users WHERE id = ?').get(result.lastInsertRowid)
    },
    async getProgress(userId) {
      return db.prepare('SELECT settings_json, courses_json, updated_at FROM progress WHERE user_id = ?').get(userId) || null
    },
    async saveProgress(userId, settings, courses) {
      db.prepare(`
        INSERT INTO progress (user_id, settings_json, courses_json, updated_at)
        VALUES (?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(user_id) DO UPDATE SET
          settings_json = excluded.settings_json,
          courses_json = excluded.courses_json,
          updated_at = CURRENT_TIMESTAMP
      `).run(userId, JSON.stringify(settings), JSON.stringify(courses))
    },
  }
}

const store = createStore()

function signToken(user) {
  return jwt.sign({ sub: user.id, email: user.email }, JWT_SECRET, { expiresIn: '30d' })
}

function publicUser(user) {
  return { id: Number(user.id), name: user.name, email: user.email }
}

function parseStoredJson(value) {
  return typeof value === 'string' ? JSON.parse(value) : value
}

async function auth(req, res, next) {
  const header = req.headers.authorization || ''
  const token = header.startsWith('Bearer ') ? header.slice(7) : ''
  if (!token) return res.status(401).json({ error: 'Session manquante.' })

  try {
    const payload = jwt.verify(token, JWT_SECRET)
    const user = await store.getUserById(payload.sub)
    if (!user) return res.status(401).json({ error: 'Utilisateur introuvable.' })
    req.user = user
    next()
  } catch {
    res.status(401).json({ error: 'Session expiree.' })
  }
}

app.get('/api/health', (req, res) => {
  res.json({ ok: true, database: store.kind })
})

app.post('/api/auth/register', async (req, res) => {
  const name = String(req.body.name || '').trim()
  const email = String(req.body.email || '').trim().toLowerCase()
  const password = String(req.body.password || '')

  if (name.length < 2) return res.status(400).json({ error: 'Nom trop court.' })
  if (!email.includes('@')) return res.status(400).json({ error: 'Email invalide.' })
  if (password.length < 6) return res.status(400).json({ error: 'Mot de passe: minimum 6 caracteres.' })

  const passwordHash = await bcrypt.hash(password, 12)
  try {
    const user = await store.createUser(name, email, passwordHash)
    res.status(201).json({ token: signToken(user), user: publicUser(user) })
  } catch (error) {
    if (String(error.message).includes('UNIQUE') || String(error.message).includes('duplicate key')) {
      return res.status(409).json({ error: 'Cet email existe deja.' })
    }
    res.status(500).json({ error: 'Creation impossible.' })
  }
})

app.post('/api/auth/login', async (req, res) => {
  const email = String(req.body.email || '').trim().toLowerCase()
  const password = String(req.body.password || '')
  const user = await store.getUserByEmail(email)
  if (!user) return res.status(401).json({ error: 'Email ou mot de passe incorrect.' })

  const ok = await bcrypt.compare(password, user.password_hash)
  if (!ok) return res.status(401).json({ error: 'Email ou mot de passe incorrect.' })

  res.json({ token: signToken(user), user: publicUser(user) })
})

app.get('/api/auth/me', auth, (req, res) => {
  res.json({ user: publicUser(req.user) })
})

app.get('/api/progress', auth, async (req, res) => {
  const row = await store.getProgress(req.user.id)
  if (!row) return res.json({ settings: null, courses: null, updatedAt: null })
  res.json({
    settings: parseStoredJson(row.settings_json),
    courses: parseStoredJson(row.courses_json),
    updatedAt: row.updated_at,
  })
})

app.put('/api/progress', auth, async (req, res) => {
  const settings = req.body.settings
  const courses = req.body.courses
  if (!settings || !Array.isArray(courses)) {
    return res.status(400).json({ error: 'Payload invalide.' })
  }

  await store.saveProgress(req.user.id, settings, courses)
  res.json({ ok: true })
})

app.use(express.static(path.join(__dirname, 'dist')))
app.get(/^\/(?!api).*/, (req, res) => {
  res.sendFile(path.join(__dirname, 'dist', 'index.html'))
})

store.init().then(() => {
  app.listen(PORT, () => {
    console.log(`API + app ready on port ${PORT}`)
    console.log(store.kind === 'postgres' ? 'Database: PostgreSQL' : `SQLite database: ${DB_PATH}`)
  })
}).catch((error) => {
  console.error('Server startup failed:', error)
  process.exit(1)
})
