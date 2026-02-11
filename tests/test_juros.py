from decimal import Decimal

from cotacoes_moedas.juros import calculate_cdi_daily_percent


def test_calculate_cdi_daily_percent_matches_hp12c_example() -> None:
    cdi = calculate_cdi_daily_percent(Decimal("15.00"))
    assert cdi == Decimal("0.0551310642")
