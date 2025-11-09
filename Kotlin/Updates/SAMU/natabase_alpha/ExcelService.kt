package com.belasoft.natabase_alpha

import android.content.Context
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.util.CellRangeAddress
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.text.SimpleDateFormat
import java.util.*

object ExcelService {

    private const val FILE_NAME = "mapa_producao.xlsx"
    private const val MAX_PRODUCAO = 5

    private val COR_TEXTO_CASTANHO = IndexedColors.BROWN.index
    private val COR_FUNDO_BRANCO = IndexedColors.WHITE.index
    private val COR_FUNDO_CASTANHO = IndexedColors.BROWN.index
    private val COR_FUNDO_CINZA_CLARO = IndexedColors.GREY_25_PERCENT.index
    private val COR_FUNDO_AMARELO = IndexedColors.LIGHT_YELLOW.index
    private val COR_BORDA = IndexedColors.BLACK.index
    private const val MAX_PRODUCAO_DINAMICA = 20

    fun carregarMapaProducao(context: Context): MapaProducao {
        val file = File(context.getExternalFilesDir(null), FILE_NAME)

        return if (file.exists()) {
            lerExcelExistente(file)
        } else {
            criarNovoMapa(context)
        }
    }

    private fun lerExcelExistente(file: File): MapaProducao {
        FileInputStream(file).use { fis ->
            val workbook = XSSFWorkbook(fis)
            val sheet = workbook.getSheetAt(0)

            val data = sheet.getRow(0)?.getCell(13)?.toString() ?: ""
            val diaSemana = sheet.getRow(1)?.getCell(13)?.toString() ?: ""

            val itens = mutableListOf<ItemProducao>()
            var currentRow = 5
            var ultimaCategoria = ""

            while (currentRow <= sheet.lastRowNum) {
                val row = sheet.getRow(currentRow)
                if (row == null) {
                    currentRow++
                    continue
                }

                val categoria = row.getCell(0)?.toString() ?: ""
                val produto = row.getCell(1)?.toString() ?: ""

                if (categoria.isNotBlank()) {
                    ultimaCategoria = categoria
                }

                if (produto.isBlank()) {
                    currentRow++
                    continue
                }

                val validade = row.getCell(2)?.toString() ?: ""

                // Ler produções dinamicamente
                val producoes = mutableListOf<ProducaoDia>()
                var coluna = 3

                while (coluna < sheet.getRow(3)?.lastCellNum ?: 0) {
                    val cellHeader = sheet.getRow(3)?.getCell(coluna)?.toString() ?: ""
                    if (cellHeader.startsWith("Produção")) {
                        val quantidade = row.getCell(coluna)?.toString()?.toIntOrNull() ?: 0
                        val hora = row.getCell(coluna + 1)?.toString() ?: ""
                        if (quantidade > 0 || hora.isNotBlank()) {
                            producoes.add(ProducaoDia(quantidade, hora))
                        }
                        coluna += 2
                    } else {
                        break
                    }
                }

                // Encontrar colunas de perdas e sobras (elas estarão nas últimas colunas)
                val ultimaColuna = sheet.getRow(3)?.lastCellNum?.toInt() ?: 0
                val perdas = row.getCell(ultimaColuna - 2)?.toString()?.toIntOrNull() ?: 0
                val sobras = row.getCell(ultimaColuna - 1)?.toString()?.toIntOrNull() ?: 0

                itens.add(ItemProducao(ultimaCategoria, produto, producoes, perdas, sobras))
                currentRow++
            }

            workbook.close()
            return MapaProducao(data, diaSemana, itens)
        }
    }

    private fun criarNovoMapa(context: Context): MapaProducao {
        val workbook = XSSFWorkbook()
        val sheet = workbook.createSheet("Planilha1")

        criarEstruturaCompleta(workbook, sheet)

        val file = File(context.getExternalFilesDir(null), FILE_NAME)
        FileOutputStream(file).use { fos ->
            workbook.write(fos)
        }
        workbook.close()

        return carregarMapaProducao(context)
    }

    private fun criarEstruturaCompleta(workbook: XSSFWorkbook, sheet: Sheet) {
        sheet.setColumnWidth(0, 25 * 256)
        sheet.setColumnWidth(1, 30 * 256)
        sheet.setColumnWidth(2, 12 * 256)

        for (i in 3..12) {
            sheet.setColumnWidth(i, 8 * 256)
        }

        sheet.setColumnWidth(13, 10 * 256)
        sheet.setColumnWidth(14, 10 * 256)

        val row1 = sheet.createRow(0)
        val cellA1 = criarCelulaComEstilo(workbook, row1, 0, COR_FUNDO_BRANCO, true, 14, COR_TEXTO_CASTANHO)
        cellA1.setCellValue("MAPA DE PRODUÇÃO VITRINA DE PASTELARIA")
        sheet.addMergedRegion(CellRangeAddress(0, 0, 0, 10))

        val cellL1 = criarCelulaComEstilo(workbook, row1, 11, COR_FUNDO_CASTANHO, false, 11, COR_FUNDO_BRANCO)
        cellL1.setCellValue("Data:")
        val cellM1 = criarCelulaComEstilo(workbook, row1, 12, COR_FUNDO_CASTANHO, false, 11, COR_FUNDO_BRANCO)
        sheet.addMergedRegion(CellRangeAddress(0, 0, 11, 12))

        criarCelulaComEstilo(workbook, row1, 13, COR_FUNDO_BRANCO, false, 11)
        criarCelulaComEstilo(workbook, row1, 14, COR_FUNDO_BRANCO, false, 11)

        val row2 = sheet.createRow(1)
        val cellL2 = criarCelulaComEstilo(workbook, row2, 11, COR_FUNDO_CASTANHO, false, 11, COR_FUNDO_BRANCO)
        cellL2.setCellValue("Dia Semana:")
        val cellM2 = criarCelulaComEstilo(workbook, row2, 12, COR_FUNDO_CASTANHO, false, 11, COR_FUNDO_BRANCO)
        sheet.addMergedRegion(CellRangeAddress(1, 1, 11, 12))

        criarCelulaComEstilo(workbook, row2, 13, COR_FUNDO_BRANCO, false, 11)
        criarCelulaComEstilo(workbook, row2, 14, COR_FUNDO_BRANCO, false, 11)

        sheet.createRow(2)

        val row4 = sheet.createRow(3)

        val cellA4 = criarCelulaComEstilo(workbook, row4, 0, COR_FUNDO_CASTANHO, true, 11, COR_FUNDO_BRANCO, false, true, true, true)
        cellA4.setCellValue("VITRINA DE PASTELARIA\n(TEMP. 15 a 20ºC)")
        sheet.addMergedRegion(CellRangeAddress(3, 3, 0, 1))

        val cellC4 = criarCelulaComEstilo(workbook, row4, 2, COR_FUNDO_CASTANHO, true, 11, COR_FUNDO_BRANCO, true, true, false, false)
        cellC4.setCellValue("Validade\nExposição")

        val producoesHeaders = listOf("Produção 1", "Produção 2", "Produção 3", "Produção 4", "Produção 5")
        for (i in producoesHeaders.indices) {
            val coluna = 3 + (i * 2)
            val cell = criarCelulaComEstilo(workbook, row4, coluna, COR_FUNDO_CASTANHO, true, 11, COR_FUNDO_BRANCO, true, true, false, false)
            cell.setCellValue(producoesHeaders[i])
            sheet.addMergedRegion(CellRangeAddress(3, 3, coluna, coluna + 1))
        }

        val cellPerdas = criarCelulaComEstilo(workbook, row4, 13, COR_FUNDO_CASTANHO, true, 11, COR_FUNDO_BRANCO, true, false, true, true)
        cellPerdas.setCellValue("PERDAS")

        val cellSobras = criarCelulaComEstilo(workbook, row4, 14, COR_FUNDO_CASTANHO, true, 11, COR_FUNDO_BRANCO, true, false, true, true)
        cellSobras.setCellValue("SOBRAS")

        val row5 = sheet.createRow(4)

        criarCelulaComEstilo(workbook, row5, 0, COR_FUNDO_CASTANHO, false, 10, COR_FUNDO_BRANCO, false, true, true, true)
        criarCelulaComEstilo(workbook, row5, 1, COR_FUNDO_CASTANHO, false, 10, COR_FUNDO_BRANCO, false, true, true, true)

        criarCelulaComEstilo(workbook, row5, 2, COR_FUNDO_CASTANHO, false, 10, COR_FUNDO_BRANCO, true, true, false, false)

        for (i in 0 until 5) {
            val colunaBase = 3 + (i * 2)
            val cellQt = criarCelulaComEstilo(workbook, row5, colunaBase, COR_FUNDO_CASTANHO, true, 10, COR_FUNDO_BRANCO, true, true, false, false)
            cellQt.setCellValue("Qt")
            val cellHr = criarCelulaComEstilo(workbook, row5, colunaBase + 1, COR_FUNDO_CASTANHO, true, 10, COR_FUNDO_BRANCO, true, true, false, false)
            cellHr.setCellValue("Hr")
        }

        criarCelulaComEstilo(workbook, row5, 13, COR_FUNDO_CASTANHO, false, 10, COR_FUNDO_BRANCO, true, false, true, true)
        criarCelulaComEstilo(workbook, row5, 14, COR_FUNDO_CASTANHO, false, 10, COR_FUNDO_BRANCO, true, false, true, true)

        criarProdutos(workbook, sheet)
    }

    private fun criarProdutos(workbook: XSSFWorkbook, sheet: Sheet) {
        val produtos = listOf(
            // PASTELARIA
            arrayOf("PASTELARIA", "BOLA DE BERLIM", "2 D"),
            arrayOf("", "BOLO CENOURA FATIA", "2 D"),
            arrayOf("", "BOLO CHOCOLATE FATIA", "2 D"),
            arrayOf("", "BOLO DE ARROZ", "24 Hrs"),
            arrayOf("", "BRIGADEIRO", "5 D"),
            arrayOf("", "FRIPANUTS CHOCOLATE", "24 Hrs"),
            arrayOf("", "FRIPANUTS SUGAR", "24 Hrs"),
            arrayOf("", "MINI CHAUSSON", "Próprio Dia*"),
            arrayOf("", "TRANÇA CREME E MAÇÃ", "Próprio Dia*"),
            arrayOf("", "MUFFIN CHOCOLATE", "2 D"),
            arrayOf("", "MUFFIN LIMÃO", "2 D"),
            arrayOf("", "MUFFIN NOZ", "2 D"),
            arrayOf("", "PAO DEUS SIMPLES", "24 Hrs"),
            arrayOf("", "PASTEL DE NATA", "Próprio Dia*"),
            arrayOf("", "QUEIJADA DE LEITE", "3 D"),
            arrayOf("", "QUEIJADA DE MARACUJÁ", "3 D"),
            arrayOf("", "QUEIJADA LARANJA", "3 D"),
            arrayOf("", "QUEIJADA FEIJÃO", "3 D"),
            arrayOf("", "SCONE SIMPLES", "24 Hrs"),
            arrayOf("", "TARTE MAÇÃ PREMIUM", "24 Hrs"),

            // CROISSANTS
            arrayOf("CROISSANTS", "CROISSANT CHOC E AVELÃ", "Próprio Dia*"),
            arrayOf("", "CROISSANT SIMPLES", "Próprio Dia*"),
            arrayOf("", "CROISSANT MULTICEREAIS", "Próprio Dia*"),
            arrayOf("", "CROISSANT MULTICEREAIS MISTO", "Próprio Dia*"),
            arrayOf("", "CROISSANT MISTO", "4 Hrs"),
            arrayOf("", "PÃO DE DEUS MISTO", "4 Hrs"),

            // SALGADOS
            arrayOf("SALGADOS", "CHAMUÇA", "12 Hrs"),
            arrayOf("", "CROQ CARNE", "12 Hrs"),
            arrayOf("", "EMPANADA CAPRESE", "Próprio Dia*"),
            arrayOf("", "EMPANADA FRANGO", "Próprio Dia*"),
            arrayOf("", "EMPANADA Q/F", "Próprio Dia*"),
            arrayOf("", "EMPANADAS DE ATUM", "Próprio Dia*"),
            arrayOf("", "FOLHADO MISTO CARNE", "Próprio Dia*"),
            arrayOf("", "NAPOLITANA MISTA", "Próprio Dia*"),
            arrayOf("", "PASTEIS BAC", "12 Hrs"),
            arrayOf("", "RISSOL CARNE", "12 Hrs"),
            arrayOf("", "RISSOL MARISCO", "12 Hrs"),

            // TOSTAS / SANDUÍCHES
            arrayOf("TOSTAS / SANDUÍCHES", "BAGUETE AMERICANA", "4 Hrs"),
            arrayOf("", "BAGUETE ATUM", "4 Hrs"),
            arrayOf("", "BAGUETE DELICIAS", "4 Hrs"),
            arrayOf("", "BAGUETE PRESUNTO QUEI", "4 Hrs"),
            arrayOf("", "BOLA PANADO DE PORCO", "4 Hrs"),
            arrayOf("", "PAO QUEIJO FRESCO", "4 Hrs"),
            arrayOf("", "PAO SALMAO FUMADO", "4 Hrs"),
            arrayOf("", "SD FRANGO COGUMELOS", "4 Hrs"),
            arrayOf("", "BAGUETE PRESUNTO", "4 Hrs"),
            arrayOf("", "BOLA 110 GRS MISTA", "4 Hrs"),
            arrayOf("", "BOLA PANADO DE PORCO (S/ Alface)", "12 Hrs"),
            arrayOf("", "TOSTA ATUM SALOIA", "4 Hrs"),
            arrayOf("", "TOSTA FRANGO SALOIA", "4 Hrs"),
            arrayOf("", "TOSTA MISTA SALOIA", "4 Hrs"),
            arrayOf("", "TOSTA PRESUNTO/QUEIJO SALOIA", "4 Hrs"),
            arrayOf("", "BOLO CACO MISTO", "4 Hrs"),
            arrayOf("", "SD AMERICANA", "4 Hrs"),

            // REGIONAIS
            arrayOf("REGIONAIS", "PÃO DE LÓ OVAR PEQ 85 GRS", "5 D"),
            arrayOf("", "OVOS MOLES UND", "10 D"),
            arrayOf("", "TARTES DE AMÊNDOA UND", "10 D"),
            arrayOf("", "TRAVESSEIRO SINTRA", "Primária"),
            arrayOf("", "PASTEIS TORRES VEDRAS UND", "10 D"),
            arrayOf("", "PASTEIS AGUEDA UND", "10 D"),
            arrayOf("", "PASTEIS VOUZELA UND", "7 D"),
            arrayOf("", "TORTA DE AZEITÃO UND", "24 Hrs"),
            arrayOf("", "QUEIJADA MADEIRENSE", "24 Hrs"),
            arrayOf("", "MALASADA CREME (FRESCO)", "Próprio Dia"),
            arrayOf("", "SALAME (FATIA)", "Próprio Dia"),

            // PÃO
            arrayOf("PÃO", "BAGUETE", "Próprio Dia"),
            arrayOf("", "BOLA LENHA", "Próprio Dia"),
            arrayOf("", "PÃO CEREAIS", "Próprio Dia"),
            arrayOf("", "PÃO RUSTICO FATIAS", "24 Hrs")
        )

        var rowIndex = 5
        var categoriaInicio = 5
        var categoriaAtual = ""

        for (produto in produtos) {
            val row = sheet.createRow(rowIndex)

            val categoria = produto[0]
            val nomeProduto = produto[1]
            val validade = produto[2]

            if (categoria.isNotBlank() && categoria != categoriaAtual) {
                if (categoriaAtual.isNotBlank() && rowIndex > categoriaInicio) {
                    sheet.addMergedRegion(CellRangeAddress(categoriaInicio, rowIndex - 1, 0, 0))
                }
                categoriaAtual = categoria
                categoriaInicio = rowIndex
            }

            val cellCategoria = criarCelulaComEstilo(workbook, row, 0, COR_FUNDO_BRANCO, false, 11)
            if (categoria.isNotBlank()) {
                cellCategoria.setCellValue(categoria)
            }

            val cellProduto = criarCelulaComEstilo(workbook, row, 1, COR_FUNDO_BRANCO, false, 11)
            cellProduto.setCellValue(nomeProduto)

            val cellValidade = criarCelulaComEstilo(workbook, row, 2, COR_FUNDO_CINZA_CLARO, false, 11)
            cellValidade.setCellValue(validade)

            // CORREÇÃO: Aplicar cores alternadas para as colunas de produção
            for (i in 0 until MAX_PRODUCAO) {
                val colunaQt = 3 + (i * 2)
                val colunaHr = colunaQt + 1

                // CORREÇÃO: Qt em cinza, Hr em branco (padrão alternado)
                val corFundoQt = COR_FUNDO_CINZA_CLARO
                val corFundoHr = COR_FUNDO_BRANCO

                criarCelulaComEstilo(workbook, row, colunaQt, corFundoQt, false, 10)
                criarCelulaComEstilo(workbook, row, colunaHr, corFundoHr, false, 10)
            }

            val cellPerdas = criarCelulaComEstilo(workbook, row, 13, COR_FUNDO_AMARELO, false, 10)

            val cellSobras = criarCelulaComEstilo(workbook, row, 14, COR_FUNDO_BRANCO, false, 10)

            rowIndex++
        }

        if (categoriaAtual.isNotBlank() && rowIndex > categoriaInicio) {
            sheet.addMergedRegion(CellRangeAddress(categoriaInicio, rowIndex - 1, 0, 0))
        }
    }

    private fun criarCelulaComEstilo(
        workbook: XSSFWorkbook,
        row: Row,
        colIndex: Int,
        corFundo: Short,
        negrito: Boolean,
        tamanhoFonte: Int,
        corTexto: Short? = null,
        bordaEsquerda: Boolean = true,
        bordaDireita: Boolean = true,
        bordaTopo: Boolean = true,
        bordaBase: Boolean = true
    ): Cell {
        val cell = row.createCell(colIndex)
        val estilo = workbook.createCellStyle()

        if (bordaEsquerda) {
            estilo.borderLeft = BorderStyle.THIN
            estilo.leftBorderColor = COR_BORDA
        }
        if (bordaDireita) {
            estilo.borderRight = BorderStyle.THIN
            estilo.rightBorderColor = COR_BORDA
        }
        if (bordaTopo) {
            estilo.borderTop = BorderStyle.THIN
            estilo.topBorderColor = COR_BORDA
        }
        if (bordaBase) {
            estilo.borderBottom = BorderStyle.THIN
            estilo.bottomBorderColor = COR_BORDA
        }

        estilo.fillForegroundColor = corFundo
        estilo.fillPattern = FillPatternType.SOLID_FOREGROUND

        estilo.alignment = HorizontalAlignment.CENTER
        estilo.verticalAlignment = VerticalAlignment.CENTER

        estilo.wrapText = true

        val fonte = workbook.createFont()
        fonte.bold = negrito
        fonte.fontHeightInPoints = tamanhoFonte.toShort()

        if (corTexto != null) {
            fonte.color = corTexto
        }

        estilo.setFont(fonte)

        cell.cellStyle = estilo
        return cell
    }

    fun salvarProducao(context: Context, mapa: MapaProducao, producaoIndex: Int): Boolean {
        val file = File(context.getExternalFilesDir(null), FILE_NAME)
        FileInputStream(file).use { fis ->
            val workbook = XSSFWorkbook(fis)
            val sheet = workbook.getSheetAt(0)

            val colunasNecessarias = calcularColunasNecessarias(mapa)
            criarColunasExtras(workbook, sheet, colunasNecessarias)

            val hoje = SimpleDateFormat("dd/MM/yyyy", Locale.getDefault()).format(Date())
            val diaSemana = SimpleDateFormat("EEEE", Locale("pt", "PT")).format(Date())

            val cellDataN = sheet.getRow(0).getCell(13) ?: sheet.getRow(0).createCell(13)
            cellDataN.setCellValue(hoje)

            val mergedRegionData = CellRangeAddress(0, 0, 13, 14)
            if (!isRegionMerged(sheet, mergedRegionData)) {
                sheet.addMergedRegion(mergedRegionData)
            }

            val cellDiaN = sheet.getRow(1).getCell(13) ?: sheet.getRow(1).createCell(13)
            cellDiaN.setCellValue(diaSemana.capitalize(Locale.getDefault()))

            val mergedRegionDia = CellRangeAddress(1, 1, 13, 14)
            if (!isRegionMerged(sheet, mergedRegionDia)) {
                sheet.addMergedRegion(mergedRegionDia)
            }

            var rowIndex = 5
            while (rowIndex <= sheet.lastRowNum) {
                val row = sheet.getRow(rowIndex)
                if (row == null) {
                    rowIndex++
                    continue
                }

                val produtoNome = row.getCell(1)?.toString() ?: ""
                if (produtoNome.isNotBlank()) {
                    val item = mapa.itens.find { it.produto == produtoNome }
                    if (item != null) {
                        // CORREÇÃO: Salvar TODAS as produções, não apenas uma
                        for (i in item.producoes.indices) {
                            val producao = item.producoes[i]
                            val colunaQt = 3 + (i * 2)
                            val colunaHr = colunaQt + 1

                            val cellQt = row.getCell(colunaQt) ?: row.createCell(colunaQt)
                            val cellHr = row.getCell(colunaHr) ?: row.createCell(colunaHr)

                            // Aplicar estilo correto antes de definir valores
                            aplicarEstiloCelulaProducao(workbook, cellQt, COR_FUNDO_CINZA_CLARO)
                            aplicarEstiloCelulaProducao(workbook, cellHr, COR_FUNDO_BRANCO)

                            if (producao.quantidade > 0) {
                                cellQt.setCellValue(producao.quantidade.toDouble())
                                cellHr.setCellValue(producao.hora)
                            } else {
                                // CORREÇÃO: Deixar vazio em vez de 0
                                cellQt.setCellValue("")
                                cellHr.setCellValue("")
                            }
                        }

                        // Salvar perdas e sobras
                        val ultimaColuna = sheet.getRow(3)?.lastCellNum?.toInt() ?: 15
                        val colunaPerdas = ultimaColuna - 2
                        val colunaSobras = ultimaColuna - 1

                        val cellPerdas = row.getCell(colunaPerdas) ?: row.createCell(colunaPerdas)
                        cellPerdas.setCellValue(item.perdas.toDouble())

                        val cellSobras = row.getCell(colunaSobras) ?: row.createCell(colunaSobras)
                        cellSobras.setCellValue(item.sobras.toDouble())
                    }
                }
                rowIndex++
            }

            FileOutputStream(file).use { fos ->
                workbook.write(fos)
            }
            workbook.close()
        }
        return true
    }

    private fun calcularColunasNecessarias(mapa: MapaProducao): Int {
        var maxProducoes = MAX_PRODUCAO
        mapa.itens.forEach { item ->
            // Contar produções com quantidade > 0
            val producoesPreenchidas = item.producoes.count { it.quantidade > 0 }
            if (producoesPreenchidas > maxProducoes) {
                maxProducoes = producoesPreenchidas
            }
        }
        return maxProducoes.coerceAtLeast(MAX_PRODUCAO)
    }

    // Nova função para criar colunas extras
    private fun criarColunasExtras(workbook: XSSFWorkbook, sheet: Sheet, colunasNecessarias: Int) {
        if (colunasNecessarias <= MAX_PRODUCAO) return

        val headerRow = sheet.getRow(3)
        val subHeaderRow = sheet.getRow(4)

        // Calcular quantas colunas extras precisamos adicionar
        val colunasExtras = colunasNecessarias - MAX_PRODUCAO

        // Verificar se as colunas extras já existem
        val existingProductions = contarProducoesExistentes(sheet)
        if (existingProductions >= colunasNecessarias) {
            return // As colunas já existem, não precisa criar novamente
        }

        // Criar novas colunas de produção
        for (i in 1..colunasExtras) {
            val producaoNumero = MAX_PRODUCAO + i
            val colunaBase = 3 + (producaoNumero - 1) * 2

            // Verificar se esta região já está mesclada antes de adicionar
            val mergedRegion = CellRangeAddress(3, 3, colunaBase, colunaBase + 1)
            if (!isRegionMerged(sheet, mergedRegion)) {
                // Header principal (linha 3)
                val cellHeader = criarCelulaComEstilo(workbook, headerRow, colunaBase, COR_FUNDO_CASTANHO, true, 11, COR_FUNDO_BRANCO, true, true, false, false)
                cellHeader.setCellValue("Produção $producaoNumero")
                sheet.addMergedRegion(mergedRegion)
            }

            // Sub-header (linha 4) - sempre criar/atualizar
            val cellQt = criarCelulaComEstilo(workbook, subHeaderRow, colunaBase, COR_FUNDO_CASTANHO, true, 10, COR_FUNDO_BRANCO, true, true, false, false)
            cellQt.setCellValue("Qt")
            val cellHr = criarCelulaComEstilo(workbook, subHeaderRow, colunaBase + 1, COR_FUNDO_CASTANHO, true, 10, COR_FUNDO_BRANCO, true, true, false, false)
            cellHr.setCellValue("Hr")

            // Ajustar largura das colunas
            sheet.setColumnWidth(colunaBase, 8 * 256)
            sheet.setColumnWidth(colunaBase + 1, 8 * 256)
        }

        // Calcular as NOVAS posições para Perdas e Sobras
        val novaColunaPerdas = 3 + (colunasNecessarias) * 2
        val novaColunaSobras = novaColunaPerdas + 1

        // Verificar se os headers de Perdas e Sobras já existem nas novas posições
        val perdasRegion = CellRangeAddress(3, 3, novaColunaPerdas, novaColunaPerdas)
        val sobrasRegion = CellRangeAddress(3, 3, novaColunaSobras, novaColunaSobras)

        if (!isRegionMerged(sheet, perdasRegion)) {
            // Header Perdas
            val cellPerdasNova = criarCelulaComEstilo(workbook, headerRow, novaColunaPerdas, COR_FUNDO_CASTANHO, true, 11, COR_FUNDO_BRANCO, true, false, true, true)
            cellPerdasNova.setCellValue("PERDAS")
        }

        if (!isRegionMerged(sheet, sobrasRegion)) {
            // Header Sobras
            val cellSobrasNova = criarCelulaComEstilo(workbook, headerRow, novaColunaSobras, COR_FUNDO_CASTANHO, true, 11, COR_FUNDO_BRANCO, true, false, true, true)
            cellSobrasNova.setCellValue("SOBRAS")
        }

        // Sub-header Perdas e Sobras - sempre criar/atualizar
        criarCelulaComEstilo(workbook, subHeaderRow, novaColunaPerdas, COR_FUNDO_CASTANHO, false, 10, COR_FUNDO_BRANCO, true, false, true, true)
        criarCelulaComEstilo(workbook, subHeaderRow, novaColunaSobras, COR_FUNDO_CASTANHO, false, 10, COR_FUNDO_BRANCO, true, false, true, true)

        // ATUALIZAÇÃO: Inicializar TODAS as células de dados com estilo correto
        var rowIndex = 5
        while (rowIndex <= sheet.lastRowNum) {
            val row = sheet.getRow(rowIndex)
            if (row != null) {
                // Para cada nova coluna de produção, criar células com estilo correto
                for (i in 1..colunasExtras) {
                    val producaoNumero = MAX_PRODUCAO + i
                    val colunaBase = 3 + (producaoNumero - 1) * 2

                    // CORREÇÃO: Manter o padrão Qt (cinza) / Hr (branco)
                    val corFundoQt = COR_FUNDO_CINZA_CLARO
                    val corFundoHr = COR_FUNDO_BRANCO

                    // Célula de quantidade - SEMPRE criar nova com estilo correto
                    val cellQt = row.createCell(colunaBase)
                    aplicarEstiloCelulaProducao(workbook, cellQt, corFundoQt)
                    cellQt.setCellValue("") // Inicializar vazia

                    // Célula de hora - SEMPRE criar nova com estilo correto
                    val cellHr = row.createCell(colunaBase + 1)
                    aplicarEstiloCelulaProducao(workbook, cellHr, corFundoHr)
                    cellHr.setCellValue("") // Inicializar vazia
                }

                // Encontrar colunas atuais de Perdas e Sobras (elas estarão nas últimas posições)
                val ultimaColuna = sheet.getRow(3)?.lastCellNum?.toInt() ?: 15
                val colunaPerdasAtual = ultimaColuna - 2
                val colunaSobrasAtual = ultimaColuna - 1

                // Ler valores atuais
                val perdasAtual = row.getCell(colunaPerdasAtual)?.numericCellValue ?: 0.0
                val sobrasAtual = row.getCell(colunaSobrasAtual)?.numericCellValue ?: 0.0

                // Criar novas células com os valores
                val cellPerdasNova = criarCelulaComEstilo(workbook, row, novaColunaPerdas, COR_FUNDO_AMARELO, false, 10)
                cellPerdasNova.setCellValue(perdasAtual)

                val cellSobrasNova = criarCelulaComEstilo(workbook, row, novaColunaSobras, COR_FUNDO_BRANCO, false, 10)
                cellSobrasNova.setCellValue(sobrasAtual)
            }
            rowIndex++
        }

        // Ajustar largura das novas colunas de Perdas e Sobras
        sheet.setColumnWidth(novaColunaPerdas, 10 * 256)
        sheet.setColumnWidth(novaColunaSobras, 10 * 256)
    }

    private fun limparCelulasProducaoExtras(sheet: Sheet, colunasExtras: Int, workbook: XSSFWorkbook) {
        var rowIndex = 5
        while (rowIndex <= sheet.lastRowNum) {
            val row = sheet.getRow(rowIndex)
            if (row != null) {
                for (i in 1..colunasExtras) {
                    val producaoNumero = MAX_PRODUCAO + i
                    val colunaBase = 3 + (producaoNumero - 1) * 2

                    // Determinar cor de fundo
                    val corFundo = if (producaoNumero % 2 == 1) {
                        COR_FUNDO_CINZA_CLARO
                    } else {
                        COR_FUNDO_BRANCO
                    }

                    // Recriar células com estilo correto
                    val cellQt = row.createCell(colunaBase)
                    aplicarEstiloCelulaProducao(workbook, cellQt, corFundo)
                    cellQt.setCellValue("")

                    val cellHr = row.createCell(colunaBase + 1)
                    aplicarEstiloCelulaProducao(workbook, cellHr, corFundo)
                    cellHr.setCellValue("")
                }
            }
            rowIndex++
        }
    }

    // NOVA FUNÇÃO: Aplicar estilo correto às células de produção
    private fun aplicarEstiloCelulaProducao(workbook: XSSFWorkbook, cell: Cell, corFundo: Short) {
        val estilo = workbook.createCellStyle()

        // Configurar bordas
        estilo.borderLeft = BorderStyle.THIN
        estilo.borderRight = BorderStyle.THIN
        estilo.borderTop = BorderStyle.THIN
        estilo.borderBottom = BorderStyle.THIN
        estilo.leftBorderColor = COR_BORDA
        estilo.rightBorderColor = COR_BORDA
        estilo.topBorderColor = COR_BORDA
        estilo.bottomBorderColor = COR_BORDA

        estilo.fillForegroundColor = corFundo
        estilo.fillPattern = FillPatternType.SOLID_FOREGROUND

        estilo.alignment = HorizontalAlignment.CENTER
        estilo.verticalAlignment = VerticalAlignment.CENTER
        estilo.wrapText = true

        // Usar fonte padrão (não negrito, tamanho 10)
        val fonte = workbook.createFont()
        fonte.fontHeightInPoints = 10
        estilo.setFont(fonte)

        cell.cellStyle = estilo
    }

    // Nova função para contar quantas produções já existem na planilha
    private fun contarProducoesExistentes(sheet: Sheet): Int {
        val headerRow = sheet.getRow(3) ?: return 0
        var count = 0
        var col = 3

        while (col < (headerRow.lastCellNum ?: 0)) {
            val cell = headerRow.getCell(col)
            if (cell != null && cell.toString().startsWith("Produção")) {
                count++
                col += 2 // Pular coluna de hora
            } else {
                col++
            }
        }
        return count
    }

    private fun isRegionMerged(sheet: Sheet, region: CellRangeAddress): Boolean {
        for (i in 0 until sheet.numMergedRegions) {
            val existingRegion = sheet.getMergedRegion(i)
            if (existingRegion.firstRow == region.firstRow &&
                existingRegion.lastRow == region.lastRow &&
                existingRegion.firstColumn == region.firstColumn &&
                existingRegion.lastColumn == region.lastColumn) {
                return true
            }
        }
        return false
    }

    fun limparProducoes(context: Context) {
        val file = File(context.getExternalFilesDir(null), FILE_NAME)
        FileInputStream(file).use { fis ->
            val workbook = XSSFWorkbook(fis)
            val sheet = workbook.getSheetAt(0)

            var rowIndex = 5
            while (rowIndex <= sheet.lastRowNum) {
                val row = sheet.getRow(rowIndex)
                if (row != null) {
                    // Limpar apenas colunas de produção (deixar perdas e sobras)
                    var col = 3
                    while (col < (sheet.getRow(3)?.lastCellNum ?: 0) - 2) { // -2 para não limpar perdas/sobras
                        val cell = row.getCell(col)
                        if (cell != null) {
                            cell.setCellValue("")
                        }
                        col++
                    }
                }
                rowIndex++
            }

            sheet.getRow(0)?.getCell(13)?.setCellValue("")
            sheet.getRow(0)?.getCell(14)?.setCellValue("")
            sheet.getRow(1)?.getCell(13)?.setCellValue("")
            sheet.getRow(1)?.getCell(14)?.setCellValue("")

            FileOutputStream(file).use { fos ->
                workbook.write(fos)
            }
            workbook.close()
        }
    }

    fun regenerarExcelCompleto(context: Context): MapaProducao {
        val file = File(context.getExternalFilesDir(null), FILE_NAME)
        if (file.exists()) {
            file.delete()
        }
        return criarNovoMapa(context)
    }
}