# Навыки для Codex

Навыки хранятся в папке `skills/`. У каждого своя папка с `SKILL.md`, справочниками и скриптами.

| Навык | Для чего | Вызов |
| --- | --- | --- |
| [fit-sheets](skills/fit-sheets/SKILL.md) | Домашние задания фитнес-курса в Excel | `$fit-sheets` |

## Установка на macOS и Linux

```sh
git clone https://github.com/lebeddasha/skills.git
cd skills
./install.sh fit-sheets
```

Команда `./install.sh` без аргументов устанавливает все навыки. По умолчанию скрипт копирует их в `~/.agents/skills`. Чтобы выбрать конкретный проект, запустите `SKILLS_DIR="/путь/к/проекту/.agents/skills" ./install.sh fit-sheets`. Если установленная версия отличается, скрипт оставит её на месте и сообщит путь.

Откройте чат Codex, приложите Excel и вызовите `$fit-sheets`. Для ручной установки, в том числе на Windows, см. [инструкцию навыка](skills/fit-sheets/INSTALL.md).

## Добавление навыков

Чтобы добавить новый навык, создайте папку `skills/имя-навыка/` и поместите в неё `SKILL.md`. Установщик подхватит её автоматически.

Проверки перед отправкой изменений:

```sh
sh -n install.sh
python3 -B -m unittest discover -s skills/fit-sheets/tests
```

История версий описана в [CHANGELOG.md](CHANGELOG.md). Репозиторий распространяется по [лицензии MIT](LICENSE).
