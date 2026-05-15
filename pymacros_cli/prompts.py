from __future__ import annotations

from collections.abc import Callable, Sequence
from typing import TypeVar


T = TypeVar("T")
MAX_SEARCH_RESULTS = 20


class CliError(Exception):
    pass


def choose_one(
    label: str,
    options: Sequence[T],
    *,
    format_option: Callable[[T], str] = str,
) -> T:
    if not options:
        raise CliError(f"No hay opciones disponibles para {label}.")

    if len(options) == 1:
        print(f"{label} {format_option(options[0])}")
        return options[0]

    print(label)
    for index, option in enumerate(options, start=1):
        print(f"  {index}. {format_option(option)}")

    choice = input("Select an option: ").strip()

    try:
        selected = int(choice)
    except ValueError as exc:
        raise CliError("La seleccion debe ser un numero.") from exc

    if selected < 1 or selected > len(options):
        raise CliError("La seleccion esta fuera de rango.")

    return options[selected - 1]


def choose_searchable(
    label: str,
    options: Sequence[T],
    *,
    format_option: Callable[[T], str] = str,
) -> T:
    if not options:
        raise CliError(f"No hay opciones disponibles para {label}.")

    print(f"{label} Type search text to filter results, or press Enter to show all.")

    while True:
        query = input("Search: ").strip()
        matches = _filter_options(options, query, format_option)

        if not matches:
            print("No matches. Try another search.")
            continue

        visible_matches = matches[:MAX_SEARCH_RESULTS]
        print(f"Found {len(matches)} match(es):")
        for index, option in enumerate(visible_matches, start=1):
            print(f"  {index}. {format_option(option)}")

        if len(matches) > MAX_SEARCH_RESULTS:
            print(f"Showing first {MAX_SEARCH_RESULTS}; refine your search for more results.")

        choice = input("Select an option, or press Enter to search again: ").strip()

        if not choice:
            continue

        try:
            selected = int(choice)
        except ValueError as exc:
            raise CliError("La seleccion debe ser un numero.") from exc

        if selected < 1 or selected > len(visible_matches):
            raise CliError("La seleccion esta fuera de rango.")

        return visible_matches[selected - 1]


def _filter_options(
    options: Sequence[T],
    query: str,
    format_option: Callable[[T], str],
) -> list[T]:
    tokens = query.lower().split()

    if not tokens:
        return list(options)

    return [
        option
        for option in options
        if all(token in format_option(option).lower() for token in tokens)
    ]


def confirm(message: str) -> bool:
    answer = input(f"{message} [y/N]: ").strip().lower()
    return answer in {"y", "yes", "s", "si"}
