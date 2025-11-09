package com.belasoft.natabase_alpha

import androidx.recyclerview.widget.RecyclerView
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.TextView
import com.belasoft.natabase_alpha.R

class CategoriaAdapter(
    private val categorias: List<String>,
    private val onCategoriaClick: (String) -> Unit
) : RecyclerView.Adapter<CategoriaAdapter.ViewHolder>() {

    class ViewHolder(view: View) : RecyclerView.ViewHolder(view) {
        val nome: TextView = view.findViewById(R.id.categoriaNome)
    }

    override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): ViewHolder {
        val view = LayoutInflater.from(parent.context)
            .inflate(R.layout.item_categoria, parent, false)
        return ViewHolder(view)
    }

    override fun onBindViewHolder(holder: ViewHolder, position: Int) {
        val categoria = categorias[position]
        holder.nome.text = categoria
        holder.itemView.setOnClickListener { onCategoriaClick(categoria) }
    }

    override fun getItemCount() = categorias.size
}