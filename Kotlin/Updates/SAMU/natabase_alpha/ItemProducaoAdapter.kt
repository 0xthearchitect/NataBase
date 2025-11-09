package com.belasoft.natabase_alpha

import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.Button
import android.widget.TextView
import android.widget.Toast
import androidx.recyclerview.widget.RecyclerView

class ItemProducaoAdapter(
    private val lista: MutableList<ItemProducao>,
    private val onQuantidadeAlterada: (ItemProducao) -> Unit,
    private val salvarCallback: (List<ItemProducao>) -> Unit,
    private val onCalculadoraRequest: (String, String) -> Unit,
    private val onConfirmarProducao: (ItemProducao, Int) -> Unit
) : RecyclerView.Adapter<ItemProducaoAdapter.ViewHolder>() {

    private val quantidadesTemporarias = mutableMapOf<String, Int>()

    class ViewHolder(view: View) : RecyclerView.ViewHolder(view) {
        val nome: TextView = view.findViewById(R.id.produtoNome)
        val tvProducoesAnteriores: TextView = view.findViewById(R.id.tvProducoesAnteriores)
        val tvQuantidadeAtual: TextView = view.findViewById(R.id.tvQuantidadeAtual)
        val btnMenos10: Button = view.findViewById(R.id.btnMenos10)
        val btnMenos1: Button = view.findViewById(R.id.btnMenos1)
        val btnMais1: Button = view.findViewById(R.id.btnMais1)
        val btnMais10: Button = view.findViewById(R.id.btnMais10)
        val btnConfirmar: Button = view.findViewById(R.id.btnConfirmarProduto)
    }

    override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): ViewHolder {
        val view = LayoutInflater.from(parent.context)
            .inflate(R.layout.item_produto, parent, false)
        return ViewHolder(view)
    }

    override fun onBindViewHolder(holder: ViewHolder, position: Int) {
        val item = lista[position]
        val produtoNome = item.produto

        holder.nome.text = item.produto

        val producoesComQuantidade = item.producoes.filter { it.quantidade > 0 }
//        if (producoesComQuantidade.isNotEmpty()) {
//            val producoesText = producoesComQuantidade.joinToString(" • ") { producao ->
//                "P${item.producoes.indexOf(producao) + 1}: ${producao.quantidade}"
//            }
//            holder.tvProducoesAnteriores.text = producoesText
//            holder.tvProducoesAnteriores.visibility = View.VISIBLE
//        } else {
//            holder.tvProducoesAnteriores.visibility = View.GONE
//        }

        val quantidadeAtual = quantidadesTemporarias[produtoNome] ?: 0
        holder.tvQuantidadeAtual.text = quantidadeAtual.toString()

        holder.btnMenos10.setOnClickListener {
            val quantidadeAtual = quantidadesTemporarias[produtoNome] ?: 0 // CORREÇÃO: Obter valor atual
            val novaQuantidade = maxOf(0, quantidadeAtual - 10)
            quantidadesTemporarias[produtoNome] = novaQuantidade
            holder.tvQuantidadeAtual.text = novaQuantidade.toString()
        }

        holder.btnMenos1.setOnClickListener {
            val quantidadeAtual = quantidadesTemporarias[produtoNome] ?: 0 // CORREÇÃO: Obter valor atual
            val novaQuantidade = maxOf(0, quantidadeAtual - 1)
            quantidadesTemporarias[produtoNome] = novaQuantidade
            holder.tvQuantidadeAtual.text = novaQuantidade.toString()
        }

        holder.btnMais1.setOnClickListener {
            val quantidadeAtual = quantidadesTemporarias[produtoNome] ?: 0 // CORREÇÃO: Obter valor atual
            val novaQuantidade = quantidadeAtual + 1
            quantidadesTemporarias[produtoNome] = novaQuantidade
            holder.tvQuantidadeAtual.text = novaQuantidade.toString()
        }

        holder.btnMais10.setOnClickListener {
            val quantidadeAtual = quantidadesTemporarias[produtoNome] ?: 0 // CORREÇÃO: Obter valor atual
            val novaQuantidade = quantidadeAtual + 10
            quantidadesTemporarias[produtoNome] = novaQuantidade
            holder.tvQuantidadeAtual.text = novaQuantidade.toString()
        }

        holder.tvQuantidadeAtual.setOnClickListener {
            onCalculadoraRequest(item.produto, "Produção")
        }

        holder.btnConfirmar.setOnClickListener {
            val quantidadeParaConfirmar = quantidadesTemporarias[produtoNome] ?: 0
            if (quantidadeParaConfirmar > 0) {
                onConfirmarProducao(item, quantidadeParaConfirmar)

                quantidadesTemporarias[produtoNome] = 0
                holder.tvQuantidadeAtual.text = "0"

                val producoesAtualizadas = item.producoes.filter { it.quantidade > 0 }
                if (producoesAtualizadas.isNotEmpty()) {
                    val producoesText = producoesAtualizadas.joinToString(" • ") { producao ->
                        "P${item.producoes.indexOf(producao) + 1}: ${producao.quantidade}"
                    }
                    holder.tvProducoesAnteriores.text = producoesText
                    holder.tvProducoesAnteriores.visibility = View.VISIBLE
                }

                Toast.makeText(holder.itemView.context,
                    "${item.produto} confirmado com $quantidadeParaConfirmar unidades",
                    Toast.LENGTH_SHORT).show()
            } else {
                Toast.makeText(holder.itemView.context,
                    "Defina uma quantidade antes de confirmar",
                    Toast.LENGTH_SHORT).show()
            }
        }
    }

    override fun getItemCount(): Int {
        return lista.size
    }

    fun atualizarQuantidade(produtoNome: String, novaQuantidade: Int) {
        quantidadesTemporarias[produtoNome] = novaQuantidade
        val itemIndex = lista.indexOfFirst { it.produto == produtoNome }
        if (itemIndex != -1) {
            notifyItemChanged(itemIndex)
        }
    }

    fun atualizarItem(produtoNome: String, tipoOperacao: String, valor: Int, producaoIndex: Int) {
        atualizarQuantidade(produtoNome, valor)
    }
}