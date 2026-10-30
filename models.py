from datetime import date, datetime, timezone
from decimal import Decimal

from sqlalchemy import (
    Boolean,
    Date,
    DateTime,
    ForeignKey,
    ForeignKeyConstraint,
    Index,
    Integer,
    JSON,
    LargeBinary,
    Numeric,
    String,
    Text,
    UniqueConstraint,
    false,
)
from sqlalchemy.orm import Mapped, mapped_column, relationship

from database import Base


def utcnow() -> datetime:
    return datetime.now(timezone.utc)


class Tenant(Base):
    __tablename__ = "tenants"

    id: Mapped[int] = mapped_column(primary_key=True)
    name: Mapped[str] = mapped_column(String(160))
    slug: Mapped[str] = mapped_column(String(180), unique=True, index=True)
    live_sending_enabled: Mapped[bool] = mapped_column(Boolean, default=False, server_default=false())
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)


class User(Base):
    __tablename__ = "users"
    __table_args__ = (
        ForeignKeyConstraint(["tenant_id"], ["tenants.id"], ondelete="CASCADE"),
        UniqueConstraint("email", name="uq_users_email"),
        UniqueConstraint("tenant_id", "id", name="uq_users_tenant_id"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    tenant_id: Mapped[int] = mapped_column(Integer, nullable=False, index=True)
    email: Mapped[str] = mapped_column(String(320), nullable=False, index=True)
    password_hash: Mapped[str] = mapped_column(String(512), nullable=False)
    role: Mapped[str] = mapped_column(String(24), default="admin")
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)


class Customer(Base):
    __tablename__ = "customers"
    __table_args__ = (
        ForeignKeyConstraint(["tenant_id"], ["tenants.id"], ondelete="CASCADE"),
        UniqueConstraint("tenant_id", "id", name="uq_customers_tenant_id"),
        Index("ix_customers_tenant_name", "tenant_id", "name"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    tenant_id: Mapped[int] = mapped_column(Integer, nullable=False, index=True)
    external_id: Mapped[str | None] = mapped_column(String(120))
    name: Mapped[str] = mapped_column(String(200), nullable=False)
    company: Mapped[str | None] = mapped_column(String(200))
    location: Mapped[str | None] = mapped_column(String(200))
    language: Mapped[str] = mapped_column(String(8), default="en")
    reminder_frequency_days: Mapped[int | None] = mapped_column(Integer)
    reminder_dates: Mapped[list[str]] = mapped_column(JSON, default=list)
    is_active: Mapped[bool] = mapped_column(Boolean, default=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)

    contacts: Mapped[list["Contact"]] = relationship(back_populates="customer", cascade="all, delete-orphan")
    invoices: Mapped[list["Invoice"]] = relationship(back_populates="customer", cascade="all, delete-orphan")


class Contact(Base):
    __tablename__ = "contacts"
    __table_args__ = (
        ForeignKeyConstraint(["tenant_id"], ["tenants.id"], ondelete="CASCADE"),
        ForeignKeyConstraint(
            ["tenant_id", "customer_id"],
            ["customers.tenant_id", "customers.id"],
            ondelete="CASCADE",
        ),
        UniqueConstraint("tenant_id", "id", name="uq_contacts_tenant_id"),
        Index("ix_contacts_customer", "tenant_id", "customer_id"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    tenant_id: Mapped[int] = mapped_column(Integer, nullable=False, index=True)
    customer_id: Mapped[int] = mapped_column(Integer, nullable=False)
    name: Mapped[str | None] = mapped_column(String(200))
    phone: Mapped[str | None] = mapped_column(String(40))
    email: Mapped[str | None] = mapped_column(String(320))
    telegram_chat_id: Mapped[str | None] = mapped_column(String(120))
    is_primary: Mapped[bool] = mapped_column(Boolean, default=False)
    sms_opted_in: Mapped[bool] = mapped_column(Boolean, default=False, server_default=false())
    whatsapp_opted_in: Mapped[bool] = mapped_column(Boolean, default=False)
    email_opted_in: Mapped[bool] = mapped_column(Boolean, default=False, server_default=false())

    customer: Mapped[Customer] = relationship(back_populates="contacts")


class Invoice(Base):
    __tablename__ = "invoices"
    __table_args__ = (
        ForeignKeyConstraint(["tenant_id"], ["tenants.id"], ondelete="CASCADE"),
        ForeignKeyConstraint(
            ["tenant_id", "customer_id"],
            ["customers.tenant_id", "customers.id"],
            ondelete="CASCADE",
        ),
        UniqueConstraint("tenant_id", "invoice_number", name="uq_invoices_tenant_number"),
        UniqueConstraint("tenant_id", "id", name="uq_invoices_tenant_id"),
        Index("ix_invoices_due_date", "tenant_id", "status", "due_date"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    tenant_id: Mapped[int] = mapped_column(Integer, nullable=False, index=True)
    customer_id: Mapped[int] = mapped_column(Integer, nullable=False)
    invoice_number: Mapped[str] = mapped_column(String(120), nullable=False)
    total_amount: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    paid_amount: Mapped[Decimal] = mapped_column(Numeric(14, 2), default=Decimal("0"))
    due_amount: Mapped[Decimal | None] = mapped_column(Numeric(14, 2))
    due_date: Mapped[datetime | None] = mapped_column(Date)
    status: Mapped[str] = mapped_column(String(24), default="open")
    document_path: Mapped[str | None] = mapped_column(String(500))
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)

    customer: Mapped[Customer] = relationship(back_populates="invoices")
    installments: Mapped[list["Installment"]] = relationship(back_populates="invoice", cascade="all, delete-orphan")


class Installment(Base):
    __tablename__ = "installments"
    __table_args__ = (
        ForeignKeyConstraint(["tenant_id"], ["tenants.id"], ondelete="CASCADE"),
        ForeignKeyConstraint(
            ["tenant_id", "invoice_id"],
            ["invoices.tenant_id", "invoices.id"],
            ondelete="CASCADE",
        ),
        UniqueConstraint("tenant_id", "invoice_id", "installment_number", name="uq_installments_sequence"),
        UniqueConstraint("tenant_id", "id", name="uq_installments_tenant_id"),
        Index("ix_installments_schedule", "tenant_id", "status", "due_date"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    tenant_id: Mapped[int] = mapped_column(Integer, nullable=False, index=True)
    invoice_id: Mapped[int] = mapped_column(Integer, nullable=False)
    installment_number: Mapped[int] = mapped_column(Integer, nullable=False)
    amount: Mapped[Decimal] = mapped_column(Numeric(14, 2), nullable=False)
    due_date: Mapped[date] = mapped_column(Date, nullable=False)
    reminder_days_before: Mapped[list[int]] = mapped_column(JSON, default=lambda: [5, 3])
    status: Mapped[str] = mapped_column(String(24), default="planned")
    paid_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))

    invoice: Mapped[Invoice] = relationship(back_populates="installments")


class ReminderPolicy(Base):
    __tablename__ = "reminder_policies"
    __table_args__ = (
        ForeignKeyConstraint(["tenant_id"], ["tenants.id"], ondelete="CASCADE"),
        ForeignKeyConstraint(
            ["tenant_id", "customer_id"],
            ["customers.tenant_id", "customers.id"],
            ondelete="CASCADE",
        ),
        UniqueConstraint("tenant_id", "scope_key", name="uq_reminder_policy_scope"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    tenant_id: Mapped[int] = mapped_column(Integer, nullable=False, index=True)
    customer_id: Mapped[int | None] = mapped_column(Integer)
    scope_key: Mapped[str] = mapped_column(String(80), default="default")
    frequency_days: Mapped[int] = mapped_column(Integer, default=3)
    max_reminders: Mapped[int] = mapped_column(Integer, default=4)
    default_reminder_days_before: Mapped[list[int]] = mapped_column(JSON, default=lambda: [3])


class ChannelSettings(Base):
    __tablename__ = "channel_settings"
    __table_args__ = (
        ForeignKeyConstraint(["tenant_id"], ["tenants.id"], ondelete="CASCADE"),
        UniqueConstraint("tenant_id", "channel", name="uq_channel_settings_tenant_channel"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    tenant_id: Mapped[int] = mapped_column(Integer, nullable=False, index=True)
    channel: Mapped[str] = mapped_column(String(24), nullable=False)
    enabled: Mapped[bool] = mapped_column(Boolean, default=False)
    provider: Mapped[str | None] = mapped_column(String(40))
    config: Mapped[dict[str, object]] = mapped_column(JSON, default=dict)
    encrypted_secrets: Mapped[bytes | None] = mapped_column(LargeBinary)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow, onupdate=utcnow)


class MessageTemplate(Base):
    __tablename__ = "message_templates"
    __table_args__ = (
        ForeignKeyConstraint(["tenant_id"], ["tenants.id"], ondelete="CASCADE"),
        UniqueConstraint("tenant_id", "channel", "language", "stage", name="uq_message_template_scope"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    tenant_id: Mapped[int] = mapped_column(Integer, nullable=False, index=True)
    channel: Mapped[str] = mapped_column(String(24), nullable=False)
    language: Mapped[str] = mapped_column(String(8), nullable=False)
    stage: Mapped[str] = mapped_column(String(40), default="reminder")
    subject: Mapped[str | None] = mapped_column(String(200))
    body: Mapped[str] = mapped_column(Text, nullable=False)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow, onupdate=utcnow)


class ReminderEvent(Base):
    __tablename__ = "reminder_events"
    __table_args__ = (
        ForeignKeyConstraint(["tenant_id"], ["tenants.id"], ondelete="CASCADE"),
        ForeignKeyConstraint(
            ["tenant_id", "invoice_id"],
            ["invoices.tenant_id", "invoices.id"],
            ondelete="CASCADE",
        ),
        ForeignKeyConstraint(
            ["tenant_id", "installment_id"],
            ["installments.tenant_id", "installments.id"],
            ondelete="CASCADE",
        ),
        UniqueConstraint("tenant_id", "idempotency_key", name="uq_reminder_event_idempotency"),
        Index("ix_reminder_events_schedule", "tenant_id", "status", "scheduled_at"),
    )

    id: Mapped[int] = mapped_column(primary_key=True)
    tenant_id: Mapped[int] = mapped_column(Integer, nullable=False, index=True)
    invoice_id: Mapped[int] = mapped_column(Integer, nullable=False)
    installment_id: Mapped[int | None] = mapped_column(Integer)
    channel: Mapped[str] = mapped_column(String(24), nullable=False)
    status: Mapped[str] = mapped_column(String(24), default="scheduled")
    scheduled_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    delivered_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    claimed_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    attempts: Mapped[int] = mapped_column(Integer, default=0, server_default="0")
    idempotency_key: Mapped[str] = mapped_column(String(180), nullable=False)
    failure_code: Mapped[str | None] = mapped_column(String(80))

    invoice: Mapped[Invoice] = relationship()
