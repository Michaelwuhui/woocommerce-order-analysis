-- PostgreSQL compatibility primitives for the legacy query surface.
-- New code uses native PostgreSQL types/functions; these helpers only keep
-- established reports and administration routes operational during migration.

SET ROLE woo_analysis_owner;
SET search_path TO public;

DO $$
BEGIN
    CREATE COLLATION nocase (
        provider = icu,
        locale = 'und-u-ks-level2',
        deterministic = false
    );
EXCEPTION
    WHEN duplicate_object THEN NULL;
END
$$;

CREATE OR REPLACE FUNCTION sqlite_datetime(VARIADIC args text[])
RETURNS text
LANGUAGE plpgsql
STABLE
PARALLEL SAFE
AS $$
DECLARE
    base text;
    modifier text;
    value_ts timestamp without time zone;
BEGIN
    base := COALESCE(args[1], 'now');
    IF lower(base) = 'now' THEN
        value_ts := timezone('UTC', CURRENT_TIMESTAMP);
    ELSE
        BEGIN
            value_ts := replace(substr(base, 1, 19), 'T', ' ')::timestamp;
        EXCEPTION WHEN others THEN
            RETURN NULL;
        END;
    END IF;

    IF COALESCE(array_length(args, 1), 0) > 1 THEN
        FOREACH modifier IN ARRAY args[2:array_length(args, 1)] LOOP
            IF modifier IS NULL OR lower(trim(modifier)) = 'utc' THEN
                CONTINUE;
            ELSIF lower(trim(modifier)) = 'localtime' THEN
                value_ts := timezone('Asia/Hong_Kong', value_ts AT TIME ZONE 'UTC');
            ELSIF lower(trim(modifier)) = 'start of day' THEN
                value_ts := date_trunc('day', value_ts);
            ELSIF lower(trim(modifier)) = 'start of month' THEN
                value_ts := date_trunc('month', value_ts);
            ELSIF trim(modifier) ~ '^[+-][0-9]+[[:space:]]+(second|seconds|minute|minutes|hour|hours|day|days|month|months|year|years)$' THEN
                value_ts := value_ts + trim(modifier)::interval;
            END IF;
        END LOOP;
    END IF;
    RETURN to_char(value_ts, 'YYYY-MM-DD HH24:MI:SS');
END
$$;

CREATE OR REPLACE FUNCTION sqlite_date(VARIADIC args text[])
RETURNS text
LANGUAGE sql
STABLE
PARALLEL SAFE
AS $$
    SELECT substr(sqlite_datetime(VARIADIC args), 1, 10)
$$;

CREATE OR REPLACE FUNCTION sqlite_strftime(format_text text, value_text text)
RETURNS text
LANGUAGE plpgsql
STABLE
PARALLEL SAFE
AS $$
DECLARE
    value_ts timestamp without time zone;
    pg_format text;
BEGIN
    IF lower(COALESCE(value_text, 'now')) = 'now' THEN
        value_ts := timezone('UTC', CURRENT_TIMESTAMP);
    ELSE
        BEGIN
            value_ts := replace(substr(value_text, 1, 19), 'T', ' ')::timestamp;
        EXCEPTION WHEN others THEN
            RETURN NULL;
        END;
    END IF;
    IF format_text = '%s' THEN
        RETURN floor(extract(epoch FROM value_ts AT TIME ZONE 'UTC'))::bigint::text;
    END IF;
    pg_format := format_text;
    pg_format := replace(pg_format, '%Y', 'YYYY');
    pg_format := replace(pg_format, '%m', 'MM');
    pg_format := replace(pg_format, '%d', 'DD');
    pg_format := replace(pg_format, '%H', 'HH24');
    pg_format := replace(pg_format, '%M', 'MI');
    pg_format := replace(pg_format, '%S', 'SS');
    pg_format := replace(pg_format, '%W', 'WW');
    pg_format := replace(pg_format, '%w', 'D');
    RETURN to_char(value_ts, pg_format);
END
$$;

CREATE OR REPLACE FUNCTION sqlite_strftime(format_text text)
RETURNS text
LANGUAGE sql
STABLE
PARALLEL SAFE
AS $$
    SELECT sqlite_strftime(format_text, 'now')
$$;

CREATE OR REPLACE FUNCTION sqlite_julianday(value_text text)
RETURNS numeric
LANGUAGE plpgsql
STABLE
PARALLEL SAFE
AS $$
DECLARE
    value_ts timestamp without time zone;
BEGIN
    IF lower(COALESCE(value_text, 'now')) = 'now' THEN
        value_ts := timezone('UTC', CURRENT_TIMESTAMP);
    ELSE
        BEGIN
            value_ts := replace(substr(value_text, 1, 19), 'T', ' ')::timestamp;
        EXCEPTION WHEN others THEN
            RETURN NULL;
        END;
    END IF;
    RETURN extract(epoch FROM value_ts AT TIME ZONE 'UTC') / 86400.0 + 2440587.5;
END
$$;

CREATE OR REPLACE FUNCTION instr(haystack text, needle text)
RETURNS integer
LANGUAGE sql
IMMUTABLE
STRICT
PARALLEL SAFE
AS $$
    SELECT strpos(haystack, needle)
$$;

CREATE OR REPLACE FUNCTION ifnull(left_value anycompatible, right_value anycompatible)
RETURNS anycompatible
LANGUAGE sql
IMMUTABLE
PARALLEL SAFE
AS $$
    SELECT COALESCE(left_value, right_value)
$$;

CREATE OR REPLACE FUNCTION json_valid(document text)
RETURNS integer
LANGUAGE plpgsql
IMMUTABLE
PARALLEL SAFE
AS $$
BEGIN
    IF document IS NULL THEN
        RETURN 0;
    END IF;
    PERFORM document::jsonb;
    RETURN 1;
EXCEPTION WHEN others THEN
    RETURN 0;
END
$$;

CREATE OR REPLACE FUNCTION json_extract(document text, json_path text)
RETURNS text
LANGUAGE plpgsql
IMMUTABLE
PARALLEL SAFE
AS $$
DECLARE
    found jsonb;
BEGIN
    IF document IS NULL OR json_path IS NULL THEN
        RETURN NULL;
    END IF;
    found := jsonb_path_query_first(document::jsonb, json_path::jsonpath);
    IF found IS NULL OR found = 'null'::jsonb THEN
        RETURN NULL;
    END IF;
    RETURN found #>> '{}';
EXCEPTION WHEN others THEN
    RETURN NULL;
END
$$;

CREATE OR REPLACE FUNCTION json_each(document text)
RETURNS TABLE(
    key text,
    value text,
    type text,
    atom text,
    id bigint,
    parent bigint,
    fullkey text,
    path text
)
LANGUAGE plpgsql
IMMUTABLE
PARALLEL SAFE
AS $$
DECLARE
    parsed jsonb;
BEGIN
    parsed := document::jsonb;
    IF jsonb_typeof(parsed) = 'object' THEN
        RETURN QUERY
        SELECT item.key,
               item.value #>> '{}',
               jsonb_typeof(item.value),
               CASE WHEN jsonb_typeof(item.value) IN ('array','object') THEN NULL ELSE item.value #>> '{}' END,
               row_number() OVER ()::bigint,
               NULL::bigint,
               '$.' || item.key,
               '$'
        FROM jsonb_each(parsed) AS item;
    ELSIF jsonb_typeof(parsed) = 'array' THEN
        RETURN QUERY
        SELECT (item.ordinality - 1)::text,
               item.value #>> '{}',
               jsonb_typeof(item.value),
               CASE WHEN jsonb_typeof(item.value) IN ('array','object') THEN NULL ELSE item.value #>> '{}' END,
               item.ordinality::bigint,
               NULL::bigint,
               '$[' || (item.ordinality - 1)::text || ']',
               '$'
        FROM jsonb_array_elements(parsed) WITH ORDINALITY AS item(value, ordinality);
    END IF;
EXCEPTION WHEN others THEN
    RETURN;
END
$$;

CREATE OR REPLACE FUNCTION group_concat_state(state text, value text)
RETURNS text
LANGUAGE sql
IMMUTABLE
PARALLEL SAFE
AS $$
    SELECT CASE
        WHEN value IS NULL THEN state
        WHEN state IS NULL OR state = '' THEN value
        ELSE state || ',' || value
    END
$$;

CREATE OR REPLACE FUNCTION group_concat_state_sep(state text, value text, separator text)
RETURNS text
LANGUAGE sql
IMMUTABLE
PARALLEL SAFE
AS $$
    SELECT CASE
        WHEN value IS NULL THEN state
        WHEN state IS NULL OR state = '' THEN value
        ELSE state || COALESCE(separator, ',') || value
    END
$$;

DROP AGGREGATE IF EXISTS group_concat(text);
CREATE AGGREGATE group_concat(text) (
    SFUNC = group_concat_state,
    STYPE = text
);

DROP AGGREGATE IF EXISTS group_concat(text, text);
CREATE AGGREGATE group_concat(text, text) (
    SFUNC = group_concat_state_sep,
    STYPE = text
);

RESET ROLE;
