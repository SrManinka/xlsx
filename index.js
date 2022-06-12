const xlsx = require("xlsx");
const puppeteer = require('puppeteer');
require('dotenv').config({ path: __dirname + '/.env' })

// id teacher
let idt = "7"
// subject position
let subjectP = 4

// note position
let np = 2

async function getInfo() {
    let fileName = "notas.xlsx"
    let path = __dirname + '/temp/' + fileName
    //read the xlsx file
    const wb = xlsx.readFile(path)
    //transforms the xlsx sheet into json
    let file = xlsx.utils.sheet_to_json(wb.Sheets["Sheet"])
    let studentList = []
    let subject
    file.map(object => {

        if (!subject)
            subject = object.Disciplina

        let student = {}
        student.name = object.Aluno
        student.grade = object["Nota do Aluno"].toFixed(1)
        // console.log(object)
        studentList.push(student)
    })
    let result = { subject, students: studentList }
    // console.log(result)
    return result
}



(async () => {

    // get data from xlsx
    let file = await getInfo()

    // for (let student of file.students)
    //     console.log(student.name)

    const browser = await puppeteer.launch({ headless: false });
    const page = await browser.newPage();
    await page.goto('https://fvc.jacad.com.br/prof/professor.logarComo.logic');

    // login
    await page.type(`#login_prof`, process.env.USER_NAME)
    await page.type(`#senha`, process.env.USER_PASSWORD)
    await page.type(`[name="usuarioProfessor.idProfessor"]`, idt)
    await page.click('[type = "submit"]')

    page.waitForSelector('#mnAcad')

    // open dropdown menu
    await page.click('#mnAcad')
    // click item from menu
    await page.click('.dropdown-menu > li:nth-child(6)');

    await page.waitForSelector('#tabela')


    // select subject from list
    const href = await page.evaluate(() => {

        return ([...document.querySelectorAll("table tbody tr td a")].map(e => e.href));
    });
    let hrefList = []
    for (let i in href) {
        if (i == 0) {
            hrefList.push(href[i])
        }
        else
            if (href[i] != href[i - 1]) {
                hrefList.push(href[i])
            }
    }
    // console.log(hrefList)

    await page.goto(hrefList[subjectP - 1])
    await page.waitForSelector('#idSubPeriodoLetivo')

    const selectElem = await page.$('#idSubPeriodoLetivo ');
    await selectElem.type('Primeiro semestre');
    await page.click('[name="filtro.ativarEdicao"]')

    await page.click('[type = "submit"]')

    await page.waitForSelector('#tb-notas')

    const studentList = await page.$$eval('#tb-notas tbody tr', rows => {
        return Array.from(rows, row => {
            const columns = row.querySelectorAll('td');
            return Array.from(columns, column => column.innerText);
        });
    });

    for (let student of file.students)
        // console.log(student.name)
        for (let i in studentList) {
            // console.log(student.name, studentList[i].name, student.name == studentList[i].name)
            if (student.name == studentList[i][0]) {
                let cellName = `[name="matriculasEdicao[${i}].notasPortal[${np - 1}].notaFormatada"]`
                await page.click(cellName)
                for (let j = 0; j < 3; j++) {
                    await page.keyboard.press('Backspace');
                }
                await page.type(cellName, student.grade);
                break
            }
        }

    // console.log(studentList); // "C2"


    //   await browser.close();
})();