#!/bin/sh
set -eu

usage() {
    printf '%s\n' 'Использование: ./install.sh [имя-навыка ...]' \
        'Без аргументов устанавливает все навыки из skills/.' \
        'Папка установки: ~/.agents/skills. Для другой папки задайте SKILLS_DIR.'
}

case "${1:-}" in
    -h|--help) usage; exit 0 ;;
esac

repo_dir=$(CDPATH= cd -P -- "$(dirname -- "$0")" && pwd)
source_dir="$repo_dir/skills"
dest_dir=${SKILLS_DIR:-"$HOME/.agents/skills"}

if [ "$#" -eq 0 ]; then
    for source in "$source_dir"/*; do
        [ -f "$source/SKILL.md" ] || continue
        set -- "$@" "${source##*/}"
    done
    if [ "$#" -eq 0 ]; then
        printf '%s\n' 'В каталоге skills/ нет навыков.' >&2
        exit 1
    fi
fi

for name do
    case "$name" in
        ''|-*|*-|*--*|*[!abcdefghijklmnopqrstuvwxyz0123456789-]*)
            printf 'Недопустимое имя навыка: %s\n' "$name" >&2
            exit 2 ;;
    esac
    if [ ! -f "$source_dir/$name/SKILL.md" ]; then
        printf 'Навык не найден: %s\n' "$name" >&2
        exit 2
    fi
done

mkdir -p -- "$dest_dir"
dest_dir=$(CDPATH= cd -P -- "$dest_dir" && pwd)
status=0
for name do
    source="$source_dir/$name"
    target="$dest_dir/$name"
    if mkdir -- "$target" 2>/dev/null; then
        trap 'rm -rf -- "$target"' EXIT
        trap 'exit 1' HUP INT TERM
        cp -R -- "$source/." "$target/"
        trap - EXIT HUP INT TERM
        printf 'Установлен %s → %s\n' "$name" "$target"
    elif [ -d "$target" ] && [ ! -L "$target" ] && diff -qr "$source" "$target" >/dev/null; then
        printf 'Уже установлен %s\n' "$name"
    else
        printf 'Не удалось установить %s. Проверьте доступ к папке или сохраните существующую версию отдельно: %s\n' "$name" "$target" >&2
        status=1
    fi
done
exit "$status"
