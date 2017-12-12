const { app, BrowserWindow } = require('electron')

let mainWindow

app.on('ready', () => {
    mainWindow = new BrowserWindow({
        width: 800,
        height: 600,
        icon: `${__dirname}/assets/img/icon.png`
    });
    mainWindow.loadURL(`file://${__dirname}/pages/main-page/main-page.html`);

})

app.on('window-all-closed', () => {
    app.quit()
})
