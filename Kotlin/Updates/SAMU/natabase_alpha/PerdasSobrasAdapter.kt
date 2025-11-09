package com.belasoft.natabase_alpha

import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.Button
import android.widget.TextView
import android.widget.Toast
import androidx.core.content.ContextCompat
import androidx.fragment.app.FragmentActivity
import androidx.recyclerview.widget.RecyclerView

class PerdasSobrasAdapter(
    private val lista: MutableList<ItemProducao>,
    private val onCalculadoraRequest: (String, String) -> Unit
) : RecyclerView.Adapter<PerdasSobrasAdapter.ViewHolder>() {

    private var context: android.content.Context? = null

    class ViewHolder(view: View) : RecyclerView.ViewHolder(view) {
        val nome: TextView = view.findViewById(R.id.produtoNome)
        val producaoInfo: TextView = view.findViewById(R.id.producaoInfo)
        val btnPerdas: Button = view.findViewById(R.id.btnPerdas)
        val btnSobras: Button = view.findViewById(R.id.btnSobras)
        val saldoInfo: TextView = view.findViewById(R.id.saldoInfo)
    }

    override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): ViewHolder {
        context = parent.context
        val view = LayoutInflater.from(parent.context)
            .inflate(R.layout.item_perdas_sobras, parent, false)
        return ViewHolder(view)
    }

    override fun onBindViewHolder(holder: ViewHolder, position: Int) {
        val item = lista[position]

        holder.nome.text = item.produto

        val producaoTotal = item.producoes.sumOf { it.quantidade }
        holder.producaoInfo.text = "Produção: $producaoTotal"

        holder.btnPerdas.text = item.perdas.toString()
        holder.btnSobras.text = item.sobras.toString()

        val saldo = producaoTotal - (item.perdas + item.sobras)
        holder.saldoInfo.text = "Saldo: $saldo"

        if (saldo < 0) {
            holder.saldoInfo.setTextColor(ContextCompat.getColor(holder.itemView.context, R.color.red))
        } else {
            holder.saldoInfo.setTextColor(ContextCompat.getColor(holder.itemView.context, R.color.deep_brown))
        }

        holder.btnPerdas.setOnClickListener {
            onCalculadoraRequest(item.produto, "Perdas")
        }

        holder.btnSobras.setOnClickListener {
            onCalculadoraRequest(item.produto, "Sobras")
        }
    }

    override fun getItemCount(): Int {
        return lista.size
    }

    private fun validarPerdasSobras(produtoNome: String, tipoOperacao: String, novoValor: Int): Boolean {
        val itemIndex = lista.indexOfFirst { it.produto == produtoNome }
        if (itemIndex == -1) return false

        val item = lista[itemIndex]
        val producaoTotal = item.producoes.sumOf { it.quantidade }

        val perdasAtuais = if (tipoOperacao == "Perdas") novoValor else item.perdas
        val sobrasAtuais = if (tipoOperacao == "Sobras") novoValor else item.sobras

        return (perdasAtuais + sobrasAtuais) <= producaoTotal
    }

    fun atualizarItem(produtoNome: String, tipoOperacao: String, valor: Int) {
        val itemIndex = lista.indexOfFirst { it.produto == produtoNome }
        if (itemIndex != -1) {
            val item = lista[itemIndex]
            val updatedItem = when (tipoOperacao) {
                "Perdas" -> item.copy(perdas = valor)
                "Sobras" -> item.copy(sobras = valor)
                else -> item
            }

            val producaoTotal = updatedItem.producoes.sumOf { it.quantidade }
            if ((updatedItem.perdas + updatedItem.sobras) <= producaoTotal) {
                lista[itemIndex] = updatedItem
                notifyItemChanged(itemIndex)
            } else {
                context?.let {
                    Toast.makeText(
                        it,
                        "Erro: Perdas + Sobras não podem exceder a produção total de $producaoTotal!",
                        Toast.LENGTH_LONG
                    ).show()
                }
            }
        }
    }

    fun getProdutosAtualizados(): List<ItemProducao> {
        return lista.toList()
    }
}