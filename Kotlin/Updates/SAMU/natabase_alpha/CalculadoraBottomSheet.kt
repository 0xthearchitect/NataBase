package com.belasoft.natabase_alpha

import android.app.Dialog
import android.content.res.Configuration
import android.os.Bundle
import android.util.DisplayMetrics
import android.view.View
import android.view.ViewGroup
import android.widget.Button
import android.widget.TextView
import androidx.core.view.updateLayoutParams
import com.google.android.material.bottomsheet.BottomSheetBehavior
import com.google.android.material.bottomsheet.BottomSheetDialog
import com.google.android.material.bottomsheet.BottomSheetDialogFragment

class CalculadoraBottomSheet : BottomSheetDialogFragment() {

    private lateinit var tvDisplay: TextView
    private lateinit var tvProdutoSelecionado: TextView
    private var currentValue = 0
    private var produtoNome: String = ""
    private var tipoOperacao: String = ""
    private var producaoIndex: Int = 0

    interface CalculadoraListener {
        fun onValorConfirmado(produtoNome: String, tipoOperacao: String, valor: Int, producaoIndex: Int)
    }

    private var listener: CalculadoraListener? = null

    companion object {
        fun newInstance(produtoNome: String, tipoOperacao: String, producaoIndex: Int): CalculadoraBottomSheet {
            val args = Bundle().apply {
                putString("PRODUTO_NOME", produtoNome)
                putString("TIPO_OPERACAO", tipoOperacao)
                putInt("PRODUCAO_INDEX", producaoIndex)
            }
            return CalculadoraBottomSheet().apply {
                arguments = args
            }
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        arguments?.let {
            produtoNome = it.getString("PRODUTO_NOME") ?: ""
            tipoOperacao = it.getString("TIPO_OPERACAO") ?: ""
            producaoIndex = it.getInt("PRODUCAO_INDEX", 0)
        }
    }

    override fun onCreateDialog(savedInstanceState: Bundle?): Dialog {
        val dialog = BottomSheetDialog(requireContext(), R.style.BottomSheetDialogTheme)
        val view = layoutInflater.inflate(R.layout.bottom_sheet_calculadora, null)
        dialog.setContentView(view)

        setupClickListeners(view)
        tvDisplay = view.findViewById(R.id.tvDisplay)
        tvProdutoSelecionado = view.findViewById(R.id.tvProdutoSelecionado)

        updateProdutoInfo()
        updateDisplay()

        dialog.setOnShowListener {
            val bottomSheetDialog = it as BottomSheetDialog
            val bottomSheet = bottomSheetDialog.findViewById<View>(com.google.android.material.R.id.design_bottom_sheet)
            bottomSheet?.let { sheet ->
                val behavior = BottomSheetBehavior.from(sheet)
                val displayMetrics = DisplayMetrics()
                requireActivity().windowManager.defaultDisplay.getMetrics(displayMetrics)

                // Verificar se é tablet na horizontal
                if (isTabletLandscape()) {
                    // Para tablet na horizontal: ocupar lado esquerdo
                    val screenWidth = displayMetrics.widthPixels
                    val desiredWidth = (screenWidth * 0.4).toInt()

                    sheet.updateLayoutParams {
                        width = desiredWidth
                        height = ViewGroup.LayoutParams.MATCH_PARENT
                    }

                    behavior.peekHeight = ViewGroup.LayoutParams.MATCH_PARENT
                    behavior.state = BottomSheetBehavior.STATE_EXPANDED
                    behavior.skipCollapsed = true

                    // Configurar para abrir do lado esquerdo
                    behavior.isFitToContents = false
                    sheet.translationX = 0f
                } else {
                    // Para outros casos: comportamento original
                    val screenHeight = displayMetrics.heightPixels
                    val desiredHeight = (screenHeight * 0.4).toInt()

                    sheet.updateLayoutParams {
                        height = desiredHeight
                        width = ViewGroup.LayoutParams.MATCH_PARENT
                    }

                    behavior.peekHeight = desiredHeight
                    behavior.state = BottomSheetBehavior.STATE_EXPANDED
                    behavior.skipCollapsed = true
                }
            }
        }

        return dialog
    }

    private fun isTabletLandscape(): Boolean {
        val configuration = resources.configuration
        return (configuration.screenLayout and Configuration.SCREENLAYOUT_SIZE_MASK) >= Configuration.SCREENLAYOUT_SIZE_LARGE &&
                configuration.orientation == Configuration.ORIENTATION_LANDSCAPE
    }

    override fun onViewCreated(view: View, savedInstanceState: Bundle?) {
        super.onViewCreated(view, savedInstanceState)
        if (!::tvDisplay.isInitialized) {
            tvDisplay = view.findViewById(R.id.tvDisplay)
            tvProdutoSelecionado = view.findViewById(R.id.tvProdutoSelecionado)
            setupClickListeners(view)
            updateProdutoInfo()
            updateDisplay()
        }
    }

    private fun updateProdutoInfo() {
        val textoOperacao = when (tipoOperacao) {
            "Produção" -> "Produção ${if (producaoIndex > 0) producaoIndex else ""}"
            "Perdas" -> "Perdas"
            "Sobras" -> "Sobras"
            else -> tipoOperacao
        }

        tvProdutoSelecionado.text = "Produto: $produtoNome - $textoOperacao"
    }

    private fun setupClickListeners(view: View) {
        view.findViewById<Button>(R.id.btn0).setOnClickListener {
            appendNumber(0)
        }
        view.findViewById<Button>(R.id.btn1).setOnClickListener {
            appendNumber(1)
        }
        view.findViewById<Button>(R.id.btn2).setOnClickListener {
            appendNumber(2)
        }
        view.findViewById<Button>(R.id.btn3).setOnClickListener {
            appendNumber(3)
        }
        view.findViewById<Button>(R.id.btn4).setOnClickListener {
            appendNumber(4)
        }
        view.findViewById<Button>(R.id.btn5).setOnClickListener {
            appendNumber(5)
        }
        view.findViewById<Button>(R.id.btn6).setOnClickListener {
            appendNumber(6)
        }
        view.findViewById<Button>(R.id.btn7).setOnClickListener {
            appendNumber(7)
        }
        view.findViewById<Button>(R.id.btn8).setOnClickListener {
            appendNumber(8)
        }
        view.findViewById<Button>(R.id.btn9).setOnClickListener {
            appendNumber(9)
        }

        view.findViewById<Button>(R.id.btnClear).setOnClickListener {
            clear()
        }
        view.findViewById<Button>(R.id.btnBackspace).setOnClickListener {
            backspace()
        }
        view.findViewById<Button>(R.id.btnConfirmar).setOnClickListener {
            confirmar()
        }
    }

    private fun appendNumber(number: Int) {
        currentValue = currentValue * 10 + number
        updateDisplay()
    }

    private fun clear() {
        currentValue = 0
        updateDisplay()
    }

    private fun backspace() {
        currentValue = currentValue / 10
        updateDisplay()
    }

    private fun confirmar() {
        listener?.onValorConfirmado(produtoNome, tipoOperacao, currentValue, producaoIndex)
        dismiss()
    }

    private fun updateDisplay() {
        tvDisplay.text = currentValue.toString()
    }

    fun setListener(listener: CalculadoraListener) {
        this.listener = listener
    }
}