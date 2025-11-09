package com.belasoft.natabase_alpha

import java.io.Serializable

data class MapaProducao(
    val data: String,
    val diaSemana: String,
    val itens: List<ItemProducao>
) : Serializable

data class ItemProducao(
    val categoria: String,
    val produto: String,
    val producoes: MutableList<ProducaoDia> = mutableListOf(),
    val perdas: Int = 0,
    val sobras: Int = 0
) : Serializable {
    val quantidade: Int
        get() = producoes.sumOf { it.quantidade }

    // Nova função para adicionar produção
    fun adicionarProducao(quantidade: Int, hora: String = "") {
        // Encontrar primeira produção vazia ou adicionar nova
        var encontrouVazia = false
        for (i in producoes.indices) {
            if (producoes[i].quantidade == 0) {
                producoes[i] = ProducaoDia(quantidade, hora)
                encontrouVazia = true
                break
            }
        }

        if (!encontrouVazia) {
            // Adicionar nova produção (sem limite fixo)
            producoes.add(ProducaoDia(quantidade, hora))
        }
    }
}

data class ProducaoDia(
    val quantidade: Int = 0,
    val hora: String = ""
) : Serializable

data class Produto(
    val nome: String,
    val tipo: String,
    var quantidade: Int = 0
) : Serializable