#!/usr/bin/env python3
"""Read-only MCP server for official MOEX ISS futures data.

The server deliberately exposes no brokerage, authentication, order, account,
or write endpoints. It uses only a fixed HTTPS host and a fixed FORTS board.
"""

from __future__ import annotations

import json
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime
from typing import Any


SERVER_NAME = "moex-futures-education"
SERVER_VERSION = "0.1.0"
PROTOCOL_VERSION = "2025-06-18"
ISS_BASE = "https://iss.moex.com/iss/engines/futures/markets/forts/boards/rfud"
ISS_DOCS = "https://www.moex.com/a2193"
DELAY_NOTICE = (
    "MOEX ISS open data can be delayed and incomplete. Use it for education and "
    "verification, not as an execution feed or a guarantee of an executable price."
)
DATA_NOTICE = (
    "Respect MOEX information-service terms and licensing. This connector does not "
    "redistribute or persist a market-data archive."
)
SECID_RE = re.compile(r"^[A-Za-z0-9_-]{1,32}$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
ALLOWED_INTERVALS = {1, 10, 60, 24}

SECURITY_COLUMNS = [
    "SECID",
    "SHORTNAME",
    "SECNAME",
    "ASSETCODE",
    "PREVSETTLEPRICE",
    "DECIMALS",
    "MINSTEP",
    "STEPPRICE",
    "LASTTRADEDATE",
    "LASTDELDATE",
    "INITIALMARGIN",
    "LOTVOLUME",
]
MARKET_COLUMNS = [
    "SECID",
    "LAST",
    "SETTLEPRICE",
    "OPENPOSITION",
    "NUMTRADES",
    "VALTODAY",
    "SYSTIME",
]


TOOLS: list[dict[str, Any]] = [
    {
        "name": "search_futures",
        "title": "Найти фьючерсы MOEX",
        "description": (
            "Ищет активные фьючерсные контракты на срочном рынке MOEX по коду, "
            "краткому или полному названию. Используйте до get_futures_contract, "
            "если точный SECID неизвестен."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "query": {
                    "type": "string",
                    "description": "Часть SECID, кода базового актива или названия.",
                    "minLength": 1,
                    "maxLength": 80,
                },
                "limit": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 30,
                    "default": 10,
                },
            },
            "required": ["query"],
            "additionalProperties": False,
        },
        "annotations": {
            "readOnlyHint": True,
            "destructiveHint": False,
            "openWorldHint": True,
            "idempotentHint": True,
        },
    },
    {
        "name": "get_futures_contract",
        "title": "Получить карточку фьючерса MOEX",
        "description": (
            "Возвращает параметры одного активного контракта и текущий отложенный "
            "рыночный снимок из официального MOEX ISS. INITIALMARGIN — текущий "
            "параметр, а не цена контракта и не максимально возможный убыток."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "secid": {
                    "type": "string",
                    "description": "Точный идентификатор контракта MOEX, например SiU6.",
                    "pattern": "^[A-Za-z0-9_-]{1,32}$",
                }
            },
            "required": ["secid"],
            "additionalProperties": False,
        },
        "annotations": {
            "readOnlyHint": True,
            "destructiveHint": False,
            "openWorldHint": True,
            "idempotentHint": True,
        },
    },
    {
        "name": "get_futures_market_snapshot",
        "title": "Получить снимок рынка фьючерсов MOEX",
        "description": (
            "Возвращает отложенный снимок рынка для списка SECID. Если список пуст, "
            "возвращает наиболее торгуемые активные контракты по обороту текущего дня."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "secids": {
                    "type": "array",
                    "items": {
                        "type": "string",
                        "pattern": "^[A-Za-z0-9_-]{1,32}$",
                    },
                    "maxItems": 20,
                    "default": [],
                },
                "limit": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 30,
                    "default": 10,
                },
            },
            "additionalProperties": False,
        },
        "annotations": {
            "readOnlyHint": True,
            "destructiveHint": False,
            "openWorldHint": True,
            "idempotentHint": True,
        },
    },
    {
        "name": "get_futures_candles",
        "title": "Получить свечи фьючерса MOEX",
        "description": (
            "Возвращает исторические свечи одного контракта из MOEX ISS для учебного "
            "анализа. Не использовать как поток котировок для реальной торговли."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "secid": {
                    "type": "string",
                    "pattern": "^[A-Za-z0-9_-]{1,32}$",
                },
                "from": {
                    "type": "string",
                    "description": "Начальная дата YYYY-MM-DD.",
                    "pattern": "^\\d{4}-\\d{2}-\\d{2}$",
                },
                "till": {
                    "type": "string",
                    "description": "Конечная дата YYYY-MM-DD.",
                    "pattern": "^\\d{4}-\\d{2}-\\d{2}$",
                },
                "interval": {
                    "type": "integer",
                    "enum": [1, 10, 60, 24],
                    "default": 60,
                    "description": "Интервал: 1/10/60 минут или 24 часа.",
                },
                "limit": {
                    "type": "integer",
                    "minimum": 1,
                    "maximum": 500,
                    "default": 100,
                },
            },
            "required": ["secid", "from", "till"],
            "additionalProperties": False,
        },
        "annotations": {
            "readOnlyHint": True,
            "destructiveHint": False,
            "openWorldHint": True,
            "idempotentHint": True,
        },
    },
]


def _now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _validate_secid(value: Any) -> str:
    if not isinstance(value, str) or SECID_RE.fullmatch(value) is None:
        raise ValueError("secid must match ^[A-Za-z0-9_-]{1,32}$")
    return value


def _bounded_int(value: Any, default: int, low: int, high: int, name: str) -> int:
    if value is None:
        return default
    if isinstance(value, bool) or not isinstance(value, int) or not low <= value <= high:
        raise ValueError(f"{name} must be an integer from {low} to {high}")
    return value


def _validate_date(value: Any, name: str) -> str:
    if not isinstance(value, str) or DATE_RE.fullmatch(value) is None:
        raise ValueError(f"{name} must use YYYY-MM-DD")
    try:
        date.fromisoformat(value)
    except ValueError as exc:
        raise ValueError(f"{name} is not a valid calendar date") from exc
    return value


def _build_url(path: str, params: dict[str, Any]) -> str:
    encoded = urllib.parse.urlencode(params, doseq=True, safe=",")
    return f"{ISS_BASE}{path}?{encoded}"


def _fetch_json(url: str) -> dict[str, Any]:
    request = urllib.request.Request(
        url,
        headers={
            "Accept": "application/json",
            "User-Agent": f"{SERVER_NAME}/{SERVER_VERSION} (education; read-only)",
        },
        method="GET",
    )
    try:
        with urllib.request.urlopen(request, timeout=20) as response:
            if response.status != 200:
                raise RuntimeError(f"MOEX ISS returned HTTP {response.status}")
            raw = response.read(8_000_000)
    except urllib.error.HTTPError as exc:
        raise RuntimeError(f"MOEX ISS returned HTTP {exc.code}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"MOEX ISS is unavailable: {exc.reason}") from exc
    try:
        payload = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError("MOEX ISS returned invalid JSON") from exc
    if not isinstance(payload, dict):
        raise RuntimeError("MOEX ISS returned an unexpected response")
    return payload


def _rows(block: Any) -> list[dict[str, Any]]:
    if not isinstance(block, dict):
        return []
    columns = block.get("columns")
    data = block.get("data")
    if not isinstance(columns, list) or not isinstance(data, list):
        return []
    return [dict(zip(columns, row)) for row in data if isinstance(row, list)]


def _metadata(source_url: str) -> dict[str, str]:
    return {
        "source": "MOEX ISS (official)",
        "source_url": source_url,
        "documentation_url": ISS_DOCS,
        "retrieved_at": _now_iso(),
        "delay_notice": DELAY_NOTICE,
        "data_usage_notice": DATA_NOTICE,
    }


def _all_contracts() -> tuple[str, list[dict[str, Any]], list[dict[str, Any]]]:
    url = _build_url(
        "/securities.json",
        {
            "iss.meta": "off",
            "iss.only": "securities,marketdata",
            "securities.columns": ",".join(SECURITY_COLUMNS),
            "marketdata.columns": ",".join(MARKET_COLUMNS),
        },
    )
    payload = _fetch_json(url)
    return url, _rows(payload.get("securities")), _rows(payload.get("marketdata"))


def _search_futures(arguments: dict[str, Any]) -> dict[str, Any]:
    query = arguments.get("query")
    if not isinstance(query, str) or not query.strip() or len(query) > 80:
        raise ValueError("query must be a non-empty string up to 80 characters")
    limit = _bounded_int(arguments.get("limit"), 10, 1, 30, "limit")
    url, securities, marketdata = _all_contracts()
    markets = {row.get("SECID"): row for row in marketdata}
    needle = query.casefold().strip()
    matches = []
    for security in securities:
        haystack = " ".join(
            str(security.get(key) or "")
            for key in ("SECID", "SHORTNAME", "SECNAME", "ASSETCODE")
        ).casefold()
        if needle in haystack:
            item = dict(security)
            item["marketdata"] = markets.get(security.get("SECID"))
            matches.append(item)
    matches.sort(
        key=lambda item: (
            not str(item.get("SECID", "")).casefold().startswith(needle),
            -((item.get("marketdata") or {}).get("VALTODAY") or 0),
            str(item.get("SECID", "")),
        )
    )
    return {
        "query": query,
        "match_count": len(matches),
        "returned_count": min(len(matches), limit),
        "contracts": matches[:limit],
        **_metadata(url),
    }


def _get_futures_contract(arguments: dict[str, Any]) -> dict[str, Any]:
    secid = _validate_secid(arguments.get("secid"))
    url = _build_url(
        f"/securities/{urllib.parse.quote(secid, safe='')}.json",
        {
            "iss.meta": "off",
            "iss.only": "securities,marketdata",
            "securities.columns": ",".join(SECURITY_COLUMNS),
            "marketdata.columns": ",".join(MARKET_COLUMNS),
        },
    )
    payload = _fetch_json(url)
    securities = _rows(payload.get("securities"))
    marketdata = _rows(payload.get("marketdata"))
    if not securities:
        raise ValueError(f"active contract {secid!r} was not found on board RFUD")
    return {
        "contract": securities[0],
        "marketdata": marketdata[0] if marketdata else None,
        "interpretation_notes": {
            "INITIALMARGIN": (
                "Current exchange parameter; not contract price, risk-budget, "
                "or maximum possible loss."
            ),
            "LAST": "Observed market value; not a guaranteed execution price.",
        },
        **_metadata(url),
    }


def _get_market_snapshot(arguments: dict[str, Any]) -> dict[str, Any]:
    limit = _bounded_int(arguments.get("limit"), 10, 1, 30, "limit")
    raw_secids = arguments.get("secids", [])
    if not isinstance(raw_secids, list) or len(raw_secids) > 20:
        raise ValueError("secids must be an array with at most 20 items")
    secids = [_validate_secid(value) for value in raw_secids]
    url, securities, marketdata = _all_contracts()
    security_map = {row.get("SECID"): row for row in securities}
    market_map = {row.get("SECID"): row for row in marketdata}
    if secids:
        selected = [secid for secid in secids if secid in security_map]
        missing = [secid for secid in secids if secid not in security_map]
    else:
        selected = [
            str(row.get("SECID"))
            for row in sorted(
                marketdata,
                key=lambda item: item.get("VALTODAY") or 0,
                reverse=True,
            )[:limit]
        ]
        missing = []
    rows = [
        {"contract": security_map.get(secid), "marketdata": market_map.get(secid)}
        for secid in selected[:limit]
    ]
    return {
        "requested_secids": secids,
        "missing_secids": missing,
        "rows": rows,
        **_metadata(url),
    }


def _get_futures_candles(arguments: dict[str, Any]) -> dict[str, Any]:
    secid = _validate_secid(arguments.get("secid"))
    from_date = _validate_date(arguments.get("from"), "from")
    till_date = _validate_date(arguments.get("till"), "till")
    if date.fromisoformat(from_date) > date.fromisoformat(till_date):
        raise ValueError("from must be on or before till")
    interval = _bounded_int(arguments.get("interval"), 60, 1, 60, "interval")
    if interval not in ALLOWED_INTERVALS:
        raise ValueError("interval must be one of 1, 10, 60, 24")
    limit = _bounded_int(arguments.get("limit"), 100, 1, 500, "limit")
    url = _build_url(
        f"/securities/{urllib.parse.quote(secid, safe='')}/candles.json",
        {
            "iss.meta": "off",
            "iss.only": "candles",
            "from": from_date,
            "till": till_date,
            "interval": interval,
        },
    )
    payload = _fetch_json(url)
    candles = _rows(payload.get("candles"))
    return {
        "secid": secid,
        "from": from_date,
        "till": till_date,
        "interval": interval,
        "available_count": len(candles),
        "returned_count": min(len(candles), limit),
        "candles": candles[:limit],
        **_metadata(url),
    }


HANDLERS = {
    "search_futures": _search_futures,
    "get_futures_contract": _get_futures_contract,
    "get_futures_market_snapshot": _get_market_snapshot,
    "get_futures_candles": _get_futures_candles,
}


def _tool_result(payload: dict[str, Any]) -> dict[str, Any]:
    return {
        "structuredContent": payload,
        "content": [
            {
                "type": "text",
                "text": json.dumps(payload, ensure_ascii=False, indent=2),
            }
        ],
        "isError": False,
    }


def _tool_error(message: str) -> dict[str, Any]:
    return {
        "content": [{"type": "text", "text": message}],
        "isError": True,
    }


def _dispatch(message: dict[str, Any]) -> dict[str, Any] | None:
    method = message.get("method")
    request_id = message.get("id")
    if method and method.startswith("notifications/"):
        return None
    if request_id is None:
        return None
    if method == "initialize":
        requested = (message.get("params") or {}).get("protocolVersion")
        protocol = requested if isinstance(requested, str) else PROTOCOL_VERSION
        result = {
            "protocolVersion": protocol,
            "capabilities": {"tools": {"listChanged": False}},
            "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
            "instructions": (
                "Read-only official MOEX ISS data for futures education. Never treat LAST "
                "as an executable quote or INITIALMARGIN as a maximum loss. No orders, "
                "accounts, or broker actions are available. Cite source_url and retrieved_at."
            ),
        }
        return {"jsonrpc": "2.0", "id": request_id, "result": result}
    if method == "ping":
        return {"jsonrpc": "2.0", "id": request_id, "result": {}}
    if method == "tools/list":
        return {"jsonrpc": "2.0", "id": request_id, "result": {"tools": TOOLS}}
    if method == "tools/call":
        params = message.get("params") or {}
        name = params.get("name")
        arguments = params.get("arguments") or {}
        handler = HANDLERS.get(name)
        if handler is None:
            result = _tool_error(f"Unknown tool: {name}")
        elif not isinstance(arguments, dict):
            result = _tool_error("Tool arguments must be an object")
        else:
            try:
                result = _tool_result(handler(arguments))
            except (ValueError, RuntimeError) as exc:
                result = _tool_error(str(exc))
            except Exception as exc:  # Last-resort boundary; details go only to stderr.
                print(f"Unexpected tool error: {exc!r}", file=sys.stderr, flush=True)
                result = _tool_error("Unexpected connector error")
        return {"jsonrpc": "2.0", "id": request_id, "result": result}
    return {
        "jsonrpc": "2.0",
        "id": request_id,
        "error": {"code": -32601, "message": f"Method not found: {method}"},
    }


def main() -> None:
    for raw_line in sys.stdin:
        line = raw_line.strip()
        if not line:
            continue
        try:
            message = json.loads(line)
            if not isinstance(message, dict):
                raise ValueError("message must be a JSON object")
            response = _dispatch(message)
        except (json.JSONDecodeError, ValueError) as exc:
            response = {
                "jsonrpc": "2.0",
                "id": None,
                "error": {"code": -32700, "message": f"Parse error: {exc}"},
            }
        if response is not None:
            sys.stdout.write(json.dumps(response, ensure_ascii=False) + "\n")
            sys.stdout.flush()


if __name__ == "__main__":
    main()
