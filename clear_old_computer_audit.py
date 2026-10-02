from database import engine
from sqlalchemy import text


with engine.begin() as connection:
    result = connection.execute(
        text("""
            DELETE FROM audit_logs
            WHERE action IN (
                'Computer created',
                'Computer deleted',
                'Computer updated'
            )
        """)
    )

    print(f"Deleted: {result.rowcount} old computer audit records.")