const GoogleSpreadsheet = require('google-spreadsheet');
const { promisify } = require('util');

const secret = require('./secret.json');

const id = '1vKSLbHslqsOJxox-ZNJx7QWlqCEI1HB2ZAWtlmmzpG0';

const accessSheet = async () => {
    try {
        const sheet = new GoogleSpreadsheet(id);
        await promisify(sheet.useServiceAccountAuth)(secret);

        const info = await promisify(sheet.getInfo)();

        const worksheet = info.worksheets[0];
        const rows = await promisify(worksheet.getRows)({

        });

        rows.forEach(row => {

            let total = parseInt(row.p1) + parseInt(row.p2) + parseInt(row.p3);
            let average = Math.ceil(total / 3);
            let absence = row.faltas;
            let limit = (25 / 100) * 60; //60 aulas total * 25% de faltas permitidas

            if (average > 70 && absence <= limit) {
                row.situação = 'Aprovada';
                row.notaparaaprovaçãofinal = 0;
                console.log("Situação alterada com sucesso!")
                row.save();
            } else if (average < 70 && average >= 50 && absence <= limit) {
                row.situação = 'Exame Final';
                let minimumGrade = 100 - average;
                row.notaparaaprovaçãofinal = minimumGrade;
                console.log("Situação alterada com sucesso!")
                row.save();
            }
            else if (average < 50 && absence <= limit) {
                row.situação = 'Reprovada por Nota';
                row.notaparaaprovaçãofinal = 0;
                console.log("Situação alterada com sucesso!")
                row.save();
            }
            else {
                row.situação = 'Reprovada por Falta';
                row.notaparaaprovaçãofinal = 0;
                console.log("Situação alterada com sucesso!")
                row.save();
            }
        })
    } catch {
        console.log("Falha");
    }
}

accessSheet()