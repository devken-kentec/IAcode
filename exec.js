const { exec } = require('child_process');

// Função para abrir um programa no Windows
function openProgram(programPath) {
    exec(`start "" "${programPath}"`, (error) => {
        if (error) {
            console.error(`Erro ao abrir o programa: ${error.message}`);
        } else {
            console.log(`Programa aberto: ${programPath}`);
        }
    });
}

// Exemplo: Substitua o caminho abaixo pelo caminho do programa que deseja abrir
const programPath = "C:\\Windows\\System32\\notepad.exe"; // Caminho do Bloco de Notas
openProgram(programPath);