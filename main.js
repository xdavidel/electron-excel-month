const { app, BrowserWindow } = require('electron')
const url = require('url')
const path = require('path');

let mainWindow

app.on('ready', () => {
    mainWindow = new BrowserWindow({});
    mainWindow.loadURL(url.format({
        pathname: path.join(__dirname, 'pages/', 'main-page/main-page.html'),
        protocol: 'file:',
        slashes: true
    }));

})

app.on('window-all-closed', () => {
    app.quit()
})