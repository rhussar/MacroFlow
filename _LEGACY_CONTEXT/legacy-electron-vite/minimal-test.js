const { app, BrowserWindow } = require('electron')

app.whenReady().then(() => {
  console.log('App is ready!')
  const win = new BrowserWindow({ width: 400, height: 300 })
  win.loadURL('data:text/html,<h1>Hello World!</h1>')

  setTimeout(() => {
    console.log('Closing app...')
    app.quit()
  }, 2000)
})
