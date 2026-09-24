-- Monthly actual-settlement rates for sales-board compensation calculations.
-- Apply before deploying code that reads this table. Existing fallback rates
-- and partner receipt records are unchanged.

SET ROLE woo_analysis_owner;
SET search_path TO public;

CREATE TABLE IF NOT EXISTS sales_board_settlement_rates (
    year_month text NOT NULL,
    currency text NOT NULL,
    rate_to_cny numeric NOT NULL CHECK (rate_to_cny > 0),
    updated_at timestamptz NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_by text NOT NULL DEFAULT '',
    PRIMARY KEY (year_month, currency)
);

GRANT SELECT, INSERT, UPDATE, DELETE
ON sales_board_settlement_rates TO woo_analysis_app;

RESET ROLE;
