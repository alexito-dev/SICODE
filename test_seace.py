import asyncio
from app.modulos.seace.seace_service import process_rucs

async def main():
    result = await process_rucs(["20549322434"], "2026-03-11")
    print(f"Exitosos: {result['success_rucs']}")
    print(f"Fallidos: {result['failed_rucs']}")
    print(f"Contratos: {len(result['results'])}")

asyncio.run(main())
