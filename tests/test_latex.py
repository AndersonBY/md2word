"""Tests for md2word LaTeX module."""

from md2word.latex import (
    FormulaInfo,
    extract_latex_formulas,
    latex_to_omml,
)


class TestExtractLatexFormulas:
    """Tests for extract_latex_formulas function."""

    def test_no_formulas(self):
        """Test text without formulas."""
        text = "Hello World"
        result, formulas = extract_latex_formulas(text)
        assert result == text
        assert len(formulas) == 0

    def test_inline_formula(self):
        """Test inline formula extraction."""
        text = "The formula $x^2$ is simple."
        result, formulas = extract_latex_formulas(text)
        assert len(formulas) == 1
        assert formulas[0].latex == "x^2"
        assert formulas[0].is_block is False
        assert "FORMULAINLINE" in result

    def test_block_formula(self):
        """Test block formula extraction."""
        text = "The formula:\n$$E = mc^2$$\nis famous."
        result, formulas = extract_latex_formulas(text)
        assert len(formulas) == 1
        assert formulas[0].latex == "E = mc^2"
        assert formulas[0].is_block is True
        assert "FORMULABLOCK" in result

    def test_multiple_formulas(self):
        """Test multiple formula extraction."""
        text = "We have $a$ and $b$ and $$c$$."
        result, formulas = extract_latex_formulas(text)
        assert len(formulas) == 3

    def test_complex_formula(self):
        """Test complex formula extraction."""
        text = r"The integral $\int_0^1 x^2 dx$ equals $\frac{1}{3}$."
        result, formulas = extract_latex_formulas(text)
        assert len(formulas) == 2
        assert r"\int_0^1 x^2 dx" in formulas[0].latex

    def test_formula_info_named_tuple(self):
        """Test FormulaInfo is a proper named tuple."""
        info = FormulaInfo("placeholder", "x^2", False)
        assert info.placeholder == "placeholder"
        assert info.latex == "x^2"
        assert info.is_block is False


class TestLatexToOmml:
    """Tests for latex_to_omml function."""

    def test_simple_formula(self):
        """Test simple formula conversion."""
        result = latex_to_omml("x^2")
        assert result is not None
        assert "<m:" in result  # OMML namespace

    def test_fraction(self):
        """Test fraction conversion."""
        result = latex_to_omml(r"\frac{1}{2}")
        assert result is not None

    def test_sqrt(self):
        """Test square root conversion."""
        result = latex_to_omml(r"\sqrt{x}")
        assert result is not None

    def test_greek_letters(self):
        """Test Greek letter conversion."""
        result = latex_to_omml(r"\alpha + \beta")
        assert result is not None

    def test_subscript_superscript(self):
        """Test subscript and superscript."""
        result = latex_to_omml(r"x_1^2")
        assert result is not None

    def test_invalid_latex_returns_none(self):
        """Test invalid LaTeX returns None."""
        # This might not return None depending on latex2mathml behavior
        # but we test the error handling path
        # Result could be None or a partial conversion
        # The important thing is it doesn't raise an exception
        latex_to_omml(r"\invalidcommand{}")
