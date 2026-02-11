import { useState, useEffect } from 'react'
import './index.css'

function App() {
    const [turma, setTurma] = useState('')
    const [alunosText, setAlunosText] = useState('')
    const [numColunas, setNumColunas] = useState(40)
    const [autoColumns, setAutoColumns] = useState(true) // New state for auto columns
    const [isGenerating, setIsGenerating] = useState(false)

    // Load from LocalStorage on mount
    useEffect(() => {
        const savedTurma = localStorage.getItem('historico_turma')
        const savedAlunos = localStorage.getItem('historico_alunos')

        if (savedTurma) setTurma(savedTurma)
        if (savedAlunos) setAlunosText(savedAlunos)
    }, [])

    // Save to LocalStorage on change
    useEffect(() => {
        localStorage.setItem('historico_turma', turma)
    }, [turma])

    useEffect(() => {
        localStorage.setItem('historico_alunos', alunosText)
    }, [alunosText])

    // Make sure ExcelJS is available via CDN
    const ExcelJS = window.ExcelJS

    const handleGenerate = async () => {
        if (!ExcelJS) {
            alert('A biblioteca ExcelJS não foi carregada corretamente via CDN. Verifique a conexão com a internet.')
            return
        }

        if (!turma.trim()) {
            alert('Por favor, informe o nome da turma.')
            return
        }

        const alunos = alunosText
            .split('\n')
            .map(line => line.trim())
            .filter(line => line.length > 0)

        if (alunos.length === 0) {
            alert('Por favor, insira pelo menos um aluno.')
            return
        }

        setIsGenerating(true)

        try {
            const workbook = new ExcelJS.Workbook()
            const worksheet = workbook.addWorksheet('Lista de Presença', {
                pageSetup: {
                    paperSize: 9, // A4
                    orientation: 'landscape',
                    margins: {
                        left: 0.25, right: 0.25, top: 0.25, bottom: 0.25,
                        header: 0, footer: 0
                    }
                }
            })

            // --- CÁLCULO DE DIMENSÕES ---
            // Em pontos (Excel usa approx 1/72 inch):
            // Height Usável: ~540 points
            // Width Usável (unidades de coluna Excel): ~160 units (variável, mas calibrado para Arial 10)

            // 1. Configurar Alturas
            // A4 Landscape Height = 595 pt
            const PAGE_HEIGHT = 595
            const MARGIN_PT = 15

            const TITLE_HEIGHT = 35
            const HEADER_HEIGHT = 80

            // Altura Disponível
            const AVAILABLE_HEIGHT_RAW = PAGE_HEIGHT - (2 * MARGIN_PT) - TITLE_HEIGHT - HEADER_HEIGHT - 10

            // 2. Configurar Colunas
            const FIXED_WIDTH = 4 + 35 // Nº + Nome = 39

            // Largura Total ALVO AGRESSIVA
            // O valor padrão de 160 parecia insuficiente visualmente.
            // Vamos aumentar o alvo para GARANTIR contato com a margem direita.
            // Se exceder um pouco, o fitToPage do Excel resolve encolhendo levemente,
            // mas é melhor exceder e encolher do que sobrar borda.
            const TARGET_TOTAL_WIDTH = 166

            // Capacidade segura para CONTAR colunas
            // Mantemos ~145 para ter menos colunas que o máximo possível
            const MAX_WIDTH_CAPACITY = 145

            let finalNumColunas = numColunas
            let dynamicDateColWidth = 3.0

            if (autoColumns) {
                // Modo Auto: Calcula número seguro de colunas
                const availableForDates = MAX_WIDTH_CAPACITY - FIXED_WIDTH
                // Usamos largura 3.2 para estimativa inicial mais espaçada
                finalNumColunas = Math.floor(availableForDates / 3.2)
            }
            // Se Auto estiver false, usamos numColunas do slider

            // Lógica Crucial: FORÇAR LARGURA TOTAL
            // Largura Desejada para Datas = Alvo Total - Fixas
            const targetDatesTotalWidth = TARGET_TOTAL_WIDTH - FIXED_WIDTH

            // Largura INDIVIDUAL = Total Disponível / Número de Colunas
            dynamicDateColWidth = targetDatesTotalWidth / finalNumColunas

            // Trava para evitar colunas estranhas
            if (dynamicDateColWidth < 2.5) dynamicDateColWidth = 2.5
            if (dynamicDateColWidth > 6.0) dynamicDateColWidth = 6.0

            // Recálculo da Largura Real Pós-Ajuste
            const realTotalWidth = FIXED_WIDTH + (finalNumColunas * dynamicDateColWidth)

            // Zoom Ratio (Compensação de Altura)
            // Calculamos quanto o Excel vai ter que reduzir (scale down) para caber na página
            // Base A4 width capacity ~115 (scale 100%)
            let zoomRatio = 1
            if (realTotalWidth > 115) {
                zoomRatio = realTotalWidth / 115
            }

            // Aumentar ratio levemente para garantir que linhas encostem no fim (fator de segurança 1.05)
            // Já que estamos alargando horizontalmente, o Excel vai reduzir verticalmente mais ainda.
            zoomRatio = zoomRatio * 1.05

            // Altura por linha COMPENSADA
            const compensatedAvailableHeight = AVAILABLE_HEIGHT_RAW * zoomRatio
            let rowHeight = compensatedAvailableHeight / alunos.length

            // Limites estéticos compensados
            if (rowHeight < (15 * zoomRatio)) rowHeight = (15 * zoomRatio)
            if (rowHeight > (60 * zoomRatio)) rowHeight = (60 * zoomRatio)

            // Fonte dinâmica
            let fontSize = 12 * Math.sqrt(zoomRatio)
            if (rowHeight / zoomRatio < 25) fontSize = 11 * Math.sqrt(zoomRatio)
            if (rowHeight / zoomRatio < 20) fontSize = 10 * Math.sqrt(zoomRatio)

            // Definição de Colunas
            const columns = [
                { header: 'Nº', key: 'no', width: 4 },
                { header: 'NOME', key: 'nome', width: 35 },
            ]

            for (let i = 1; i <= finalNumColunas; i++) {
                columns.push({ header: '', key: `d${i}`, width: dynamicDateColWidth })
            }

            worksheet.columns = columns

            // --- CONSTRUÇÃO DA PLANILHA ---

            // LINHA 1: TÍTULO
            worksheet.insertRow(1, [turma.toUpperCase()])
            const lastColLetter = worksheet.getColumn(columns.length).letter
            worksheet.mergeCells(`A1:${lastColLetter}1`)

            const titleRow = worksheet.getRow(1)
            titleRow.height = TITLE_HEIGHT
            titleRow.font = { name: 'Arial', size: 16, bold: true }
            titleRow.alignment = { vertical: 'middle', horizontal: 'center' }

            // LINHA 2: CABEÇALHO
            const headerRow = worksheet.getRow(2)
            const headerValues = ['Nº', 'NOME']
            for (let i = 0; i < finalNumColunas; i++) headerValues.push('')
            headerRow.values = headerValues
            headerRow.height = HEADER_HEIGHT

            headerRow.eachCell((cell, colNumber) => {
                cell.font = { name: 'Arial', size: 10, bold: true }
                cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }

                if (colNumber <= 2) {
                    cell.alignment = { vertical: 'middle', horizontal: colNumber === 1 ? 'center' : 'left', indent: colNumber === 2 ? 1 : 0 }
                } else {
                    cell.alignment = { textRotation: 90, vertical: 'bottom', horizontal: 'center' }
                    const cycle = (colNumber - 3) % 3
                    const argbColor = cycle === 0 ? 'FFEBFFEB' : (cycle === 1 ? 'FFFFF9C4' : 'FFFFFFFF')
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argbColor } }
                }
            })

            // LINHAS: ALUNOS
            alunos.forEach((aluno, index) => {
                const rowValues = [index + 1, aluno.toUpperCase()]
                for (let i = 0; i < finalNumColunas; i++) rowValues.push('')

                const row = worksheet.addRow(rowValues)
                row.height = rowHeight

                row.eachCell((cell, colNumber) => {
                    cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } }
                    cell.font = { name: 'Arial', size: fontSize }
                    cell.alignment = {
                        vertical: 'middle',
                        horizontal: colNumber === 2 ? 'left' : 'center',
                        indent: colNumber === 2 ? 1 : 0
                    }
                })
            })

            // Ajuste final de impressão
            worksheet.pageSetup.printArea = `A1:${lastColLetter}${alunos.length + 2}`
            // Se usamos fitToPage com 1x1, o Excel força caber. 
            // Como calculamos as dimensões manualmente para "encher" a folha, 
            // o fitToWi/He pode acabar reduzindo se nosso calculo passou um pouco.
            // Vamos deixar fit ativo mas relaxado se o calculo estiver bom.
            // Para garantir "extritamente uma pagina", mantemos fit:
            worksheet.pageSetup.fitToPage = true
            worksheet.pageSetup.fitToWidth = 1
            worksheet.pageSetup.fitToHeight = 1

            const buffer = await workbook.xlsx.writeBuffer()
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
            const url = window.URL.createObjectURL(blob)
            const anchor = document.createElement('a')
            anchor.href = url
            const dateStr = new Date().toISOString().split('T')[0]
            const safeTurma = turma.replace(/[^a-z0-9]/gi, '_').toLowerCase() || 'turma'
            anchor.download = `Lista_Presenca_${safeTurma}_${dateStr}.xlsx`
            anchor.click()
            window.URL.revokeObjectURL(url)

        } catch (error) {
            console.error('Erro ao gerar Excel:', error)
            alert('Ocorreu um erro ao gerar o arquivo. Verifique o console.')
        } finally {
            setIsGenerating(false)
        }
    }

    return (
        <div className="container">
            <h1>Gerador de Lista de Presença</h1>
            <p className="subtitle">Crie folhas de chamada profissionais em segundos.</p>

            <div className="card dashboard-layout">
                <div className="content-grid">
                    {/* Left Column: Configuration */}
                    <div className="config-section">
                        <div className="form-group">
                            <label htmlFor="turma">Nome da Turma</label>
                            <input
                                id="turma"
                                type="text"
                                placeholder="Ex: 3º Ano B"
                                value={turma}
                                onChange={(e) => setTurma(e.target.value)}
                            />
                        </div>

                        <div className="form-group">
                            <label htmlFor="colunas">Colunas de Aula</label>
                            <div className="flex-row">
                                <input
                                    id="colunas"
                                    type="range"
                                    min="10"
                                    max="60"
                                    value={numColunas}
                                    className="flex-1"
                                    onChange={(e) => {
                                        setNumColunas(parseInt(e.target.value))
                                        setAutoColumns(false) // Auto-disable auto mode when sliding
                                    }}
                                    aria-label="Quantidade de colunas de aula"
                                    style={{ cursor: 'pointer' }}
                                />
                                <span style={{ minWidth: '2rem', fontWeight: 'bold', textAlign: 'center' }}>{numColunas}</span>
                            </div>
                        </div>

                        <div className="form-group checkbox-group">
                            <input
                                type="checkbox"
                                id="autoColumns"
                                checked={autoColumns}
                                onChange={(e) => setAutoColumns(e.target.checked)}
                                aria-label="Ativar ajuste automático de colunas"
                            />
                            <label htmlFor="autoColumns">
                                Automático (Maximizar)
                            </label>
                        </div>

                        <button
                            className="btn-primary"
                            onClick={handleGenerate}
                            disabled={isGenerating}
                        >
                            {isGenerating ? 'Gerando...' : 'Baixar Lista Excel'}
                        </button>
                    </div>

                    {/* Right Column: Data */}
                    <div className="list-section">
                        <label htmlFor="alunos">
                            Lista de Alunos
                            <span className="hint" style={{ color: '#9ca3af', display: 'inline', marginLeft: '0.5rem' }}>
                                (Um por linha)
                            </span>
                        </label>
                        <textarea
                            id="alunos"
                            placeholder="Ana Silva&#10;Bruno Santos..."
                            value={alunosText}
                            onChange={(e) => setAlunosText(e.target.value)}
                            aria-label="Lista de nomes dos alunos"
                        />
                    </div>
                </div>
            </div>

            <p style={{ textAlign: 'center', color: '#9ca3af', fontSize: '0.8rem', marginTop: '2rem' }}>
                Feito com ❤️ e React + ExcelJS
            </p>
        </div>
    )
}

export default App
