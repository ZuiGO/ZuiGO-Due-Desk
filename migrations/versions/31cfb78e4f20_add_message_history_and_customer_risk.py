"""add message history and customer risk flags

Revision ID: 31cfb78e4f20
Revises: daa42350065b
Create Date: 2026-10-31 10:45:00.000000

"""
from datetime import datetime, timezone
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa


revision: str = "31cfb78e4f20"
down_revision: Union[str, Sequence[str], None] = "daa42350065b"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    op.add_column(
        "customers",
        sa.Column("risk_flagged", sa.Boolean(), nullable=False, server_default=sa.false()),
    )
    op.add_column("customers", sa.Column("risk_flag_reason", sa.String(length=500), nullable=True))
    op.add_column("customers", sa.Column("risk_flagged_at", sa.DateTime(timezone=True), nullable=True))
    op.create_table(
        "message_logs",
        sa.Column("id", sa.Integer(), nullable=False),
        sa.Column("tenant_id", sa.Integer(), nullable=False),
        sa.Column("customer_id", sa.Integer(), nullable=False),
        sa.Column("contact_id", sa.Integer(), nullable=True),
        sa.Column("reminder_event_id", sa.Integer(), nullable=True),
        sa.Column("channel", sa.String(length=24), nullable=False),
        sa.Column("direction", sa.String(length=24), nullable=False),
        sa.Column("status", sa.String(length=24), nullable=False),
        sa.Column("recipient", sa.String(length=320), nullable=True),
        sa.Column("subject", sa.String(length=200), nullable=True),
        sa.Column("body", sa.Text(), nullable=True),
        sa.Column("failure_code", sa.String(length=200), nullable=True),
        sa.Column("sent_at", sa.DateTime(timezone=True), nullable=True),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["tenant_id"], ["tenants.id"], ondelete="CASCADE"),
        sa.ForeignKeyConstraint(
            ["tenant_id", "customer_id"], ["customers.tenant_id", "customers.id"], ondelete="CASCADE"
        ),
        sa.ForeignKeyConstraint(["contact_id"], ["contacts.id"], ondelete="SET NULL"),
        sa.ForeignKeyConstraint(["reminder_event_id"], ["reminder_events.id"], ondelete="SET NULL"),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_message_logs_tenant_created", "message_logs", ["tenant_id", "created_at"])
    op.create_index(op.f("ix_message_logs_tenant_id"), "message_logs", ["tenant_id"])

    templates = sa.table(
        "message_templates",
        sa.column("tenant_id", sa.Integer()),
        sa.column("channel", sa.String()),
        sa.column("language", sa.String()),
        sa.column("stage", sa.String()),
        sa.column("subject", sa.String()),
        sa.column("body", sa.Text()),
        sa.column("updated_at", sa.DateTime(timezone=True)),
    )
    starters = {
        "sms": {
            "reminder": "Hello {name}, invoice {invoice} has {amount} due on {due_date}. Please arrange payment.",
            "installment": "Hello {name}, installment of {amount} for invoice {invoice} is due on {installment_date}.",
        },
        "whatsapp": {
            "reminder": "Hello {name}, your balance of {amount} for invoice {invoice} is due on {due_date}.",
            "installment": "Hello {name}, installment of {amount} for invoice {invoice} is due on {installment_date}.",
        },
        "email": {
            "reminder": "Hello {name}, invoice {invoice} has {amount} due on {due_date}. Please arrange payment.",
            "installment": "Hello {name}, installment of {amount} for invoice {invoice} is due on {installment_date}.",
        },
        "telegram": {
            "reminder": "Hello {name}, invoice {invoice} has {amount} due on {due_date}. Please arrange payment.",
            "installment": "Hello {name}, installment of {amount} for invoice {invoice} is due on {installment_date}.",
        },
    }
    now = datetime.now(timezone.utc)
    bind = op.get_bind()
    tenant_ids = bind.execute(sa.text("SELECT id FROM tenants")).scalars().all()
    existing = {
        (tenant_id, channel, language, stage)
        for tenant_id, channel, language, stage in bind.execute(
            sa.text("SELECT tenant_id, channel, language, stage FROM message_templates")
        ).all()
    }
    rows = [
        {
            "tenant_id": tenant_id,
            "channel": channel,
            "language": "en",
            "stage": stage,
            "subject": "Payment reminder" if channel == "email" else None,
            "body": body,
            "updated_at": now,
        }
        for tenant_id in tenant_ids
        for channel, stages in starters.items()
        for stage, body in stages.items()
        if (tenant_id, channel, "en", stage) not in existing
    ]
    if rows:
        op.bulk_insert(templates, rows)


def downgrade() -> None:
    op.drop_index(op.f("ix_message_logs_tenant_id"), table_name="message_logs")
    op.drop_index("ix_message_logs_tenant_created", table_name="message_logs")
    op.drop_table("message_logs")
    op.drop_column("customers", "risk_flagged_at")
    op.drop_column("customers", "risk_flag_reason")
    op.drop_column("customers", "risk_flagged")
