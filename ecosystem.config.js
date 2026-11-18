module.exports = {
  apps: [{
    name: "51la-daily",
    script: "main.py",
    interpreter: "/usr/bin/python3",
    cwd: __dirname,
    args: "--schedule-mode --method gmail --lang zh --precreated-monthly --allow-real-email"
  }]
}