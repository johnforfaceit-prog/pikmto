const express = require('express');
const router = express.Router();

router.post('/login', (req, res) => {
  const { login, password } = req.body;
  const correctLogin    = process.env.APP_LOGIN    || 'admin';
  const correctPassword = process.env.APP_PASSWORD || 'pik2026';

  if (login === correctLogin && password === correctPassword) {
    req.session.user = { login };
    return res.json({ ok: true });
  }
  res.status(401).json({ error: 'Неверный логин или пароль' });
});

router.post('/logout', (req, res) => {
  req.session.destroy();
  res.json({ ok: true });
});

module.exports = router;
