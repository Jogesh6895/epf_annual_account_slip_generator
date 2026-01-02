"""Test suite for EPF Calculator."""

import pytest
from epf_calculator import (
    calculate_contributions,
    calculate_monthly_balances,
    EPFCalculatorError,
    FileLoadError,
    SheetNotFoundError,
    DataValidationError,
)


class TestCalculateContributions:
    """Test EPF contribution calculations."""

    def test_employee_contribution(self):
        """Test employee contribution is 12% of wage."""
        ee, er, eps = calculate_contributions(10000)
        assert ee == 1200

    def test_eps_contribution(self):
        """Test EPS contribution is 8.33% of wage."""
        ee, er, eps = calculate_contributions(10000)
        assert eps == 833

    def test_employer_contribution(self):
        """Test employer contribution is 3.67% of wage."""
        ee, er, eps = calculate_contributions(10000)
        assert er == 367

    def test_total_ee_contribution(self):
        """Test total equals 12% of wage."""
        ee, er, eps = calculate_contributions(10000)
        assert ee == er + eps

    def test_rounding(self):
        """Test values are properly rounded."""
        ee, er, eps = calculate_contributions(10001)
        assert ee == 1200
        assert eps == 833
        assert er == 367

    def test_zero_wage(self):
        """Test calculation with zero wage."""
        ee, er, eps = calculate_contributions(0)
        assert ee == 0
        assert er == 0
        assert eps == 0


class TestCalculateMonthlyBalances:
    """Test monthly balance and interest calculations."""

    def test_no_withdrawals(self):
        """Test balance calculation without withdrawals."""
        ee_balances, er_balances, ee_int, er_int = calculate_monthly_balances(
            ob_ee=10000,
            ob_er=3000,
            ee_contributions=[1000, 1000],
            er_contributions=[300, 300],
            ee_withdrawals=[0, 0],
            er_withdrawals=[0, 0],
            rate=8.5,
        )
        assert ee_balances[-1] == 11000
        assert er_balances[-1] == 3300

    def test_with_withdrawals(self):
        """Test balance calculation with withdrawals."""
        ee_balances, er_balances, ee_int, er_int = calculate_monthly_balances(
            ob_ee=10000,
            ob_er=3000,
            ee_contributions=[1000, 1000],
            er_contributions=[300, 300],
            ee_withdrawals=[500, 0],
            er_withdrawals=[150, 0],
            rate=8.5,
        )
        assert ee_balances[-1] == 10500
        assert er_balances[-1] == 3150

    def test_interest_calculation(self):
        """Test interest is calculated correctly."""
        ee_balances, er_balances, ee_int, er_int = calculate_monthly_balances(
            ob_ee=12000,
            ob_er=3600,
            ee_contributions=[1000],
            er_contributions=[300],
            ee_withdrawals=[0],
            er_withdrawals=[0],
            rate=8.5,
        )
        assert ee_balances[-1] == 11000
        assert er_balances[-1] == 3300

    def test_with_withdrawals(self):
        """Test balance calculation with withdrawals."""
        ee_balances, er_balances, ee_int, er_int = calculate_monthly_balances(
            ob_ee=10000,
            ob_er=3000,
            ee_contribs=[1000, 1000],
            er_contribs=[300, 300],
            ee_withdrawals=[500, 0],
            er_withdrawals=[150, 0],
            rate=8.5,
        )
        assert ee_balances[-1] == 10500
        assert er_balances[-1] == 3150

    def test_interest_calculation(self):
        """Test interest is calculated correctly."""
        ee_balances, er_balances, ee_int, er_int = calculate_monthly_balances(
            ob_ee=12000,
            ob_er=3600,
            ee_contribs=[1000],
            er_contribs=[300],
            ee_withdrawals=[0],
            er_withdrawals=[0],
            rate=8.5,
        )
        # Average balance: (12000 + 13000) / 2 = 12500
        # Interest: (12500 * 8.5) / 1200 = 88.54 -> 89
        assert ee_int == 89
        assert er_int == 26


class TestExceptions:
    """Test custom exceptions."""

    def test_epf_calculator_error_exists(self):
        """Test base exception class exists."""
        assert issubclass(EPFCalculatorError, Exception)

    def test_file_load_error_exists(self):
        """Test FileLoadError exists."""
        assert issubclass(FileLoadError, EPFCalculatorError)

    def test_sheet_not_found_error_exists(self):
        """Test SheetNotFoundError exists."""
        assert issubclass(SheetNotFoundError, EPFCalculatorError)

    def test_data_validation_error_exists(self):
        """Test DataValidationError exists."""
        assert issubclass(DataValidationError, EPFCalculatorError)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
