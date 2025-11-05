from pathlib import Path

ENV_PATH = Path(__file__).parent / ".env"

if not ENV_PATH.exists():
    print("¡Bienvenido al Buscador de Facturas!")
    print("Necesitas un token de Microsoft Graph.")
    print("\nPega tu token aquí (puedes copiarlo de Graph Explorer):")
    token = input("> ").strip()
    
    if not token.startswith("eyJ"):
        print("Token inválido. Debe empezar con 'eyJ'")
        input("Presiona Enter para cerrar...")
        exit()
    
    # Guardar en .env
    with open(ENV_PATH, "w", encoding="utf-8") as f:
        f.write(f"GRAPH_TOKEN={token}\n")
    
    print(f"\nToken guardado en: {ENV_PATH}")
    print("¡Listo! Ya puedes usar la app.")
    input("Presiona Enter para continuar...")
else:
    print(".env encontrado. Usando token guardado.")