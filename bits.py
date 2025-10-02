import shlex
from snippets import (
    fuzzy_command, COMMANDS, print_help, open_app,
    close_windows, minimize_windows, maximize_windows,
    list_windows, shutdown_system, restart_system,
    lock_system, sleep_system, show_aliases,
    add_alias, remove_alias
)
from Welcome import welcome_banner

def repl():
    welcome_banner()
    try:
        while True:
            raw = input("bits> ").strip()
            if not raw:
                continue
            parts = shlex.split(raw)
            cmd_raw = parts[0]
            cmd = fuzzy_command(cmd_raw) or cmd_raw

            if cmd not in COMMANDS:
                print(f"Unknown command: {cmd_raw}. Type 'help'.")
                continue

            if cmd == "help":
                print_help()
            elif cmd == "exit":
                print("Goodbye.")
                break
            elif cmd == "open" and len(parts) > 1:
                open_app(" ".join(parts[1:]))
            elif cmd == "close" and len(parts) > 1:
                close_windows(" ".join(parts[1:]), all_matches=False)
            elif cmd == "closeall" and len(parts) > 1:
                close_windows(" ".join(parts[1:]), all_matches=True)
            elif cmd == "minimize" and len(parts) > 1:
                minimize_windows(" ".join(parts[1:]))
            elif cmd == "maximize" and len(parts) > 1:
                maximize_windows(" ".join(parts[1:]))
            elif cmd == "list":
                for i, t in enumerate(list_windows(), 1):
                    print(f"{i}. {t}")
            elif cmd == "shutdown":
                shutdown_system()
            elif cmd == "restart":
                restart_system()
            elif cmd == "lock":
                lock_system()
            elif cmd == "sleep":
                sleep_system()
            elif cmd == "aliases":
                show_aliases()
            elif cmd == "addalias" and len(parts) > 2:
                add_alias(parts[1], " ".join(parts[2:]))
            elif cmd == "removealias" and len(parts) > 1:
                remove_alias(parts[1])
            else:
                print("Invalid usage. Type 'help' for commands.")
    except KeyboardInterrupt:
        print("\nExiting Bits.")

if __name__ == "__main__":
    repl()
