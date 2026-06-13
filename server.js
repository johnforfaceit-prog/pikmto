require('dotenv').config();
const express = require('express');
const session = require('express-session');
const path = require('path');

const authRouter = require('./routes/auth');
const apiRouter  = require('./routes/api');
const { requireAuth } = require('./middleware/auth');

const app = express();
const PORT = process.env.PORT || 3000;

// ── Middleware ──
app.use(express.json({ limit: '20mb' }));
app.use(express.urlencoded({ extended: true }));
app.use(session({
  secret: process.env.SESSION_SECRET || 'pik-secret',
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 8 * 60 * 60 * 1000 } // 8 часов
}));

// ── Статика (только для авторизованных) ──
app.use('/css', express.static(path.join(__dirname, 'public/css')));
app.use('/js',  express.static(path.join(__dirname, 'public/js')));

// ── Роуты ──
app.use('/auth', authRouter);
app.use('/api',  requireAuth, apiRouter);

// ── Страницы ──
app.get('/login', (req, res) => {
  if (req.session.user) return res.redirect('/');
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

app.get('/', requireAuth, (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'app.html'));
});

app.get('/logout', (req, res) => {
  req.session.destroy();
  res.redirect('/login');
});

// ── Запуск ──
app.listen(PORT, () => {
  console.log('');
  console.log(`  ✓ ПИК-сервер запущен → http://localhost:${PORT}`);
  console.log(`  Логин: ${process.env.APP_LOGIN || 'admin'}`);
  console.log('');
});
