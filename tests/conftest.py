import pytest
from fastapi.testclient import TestClient
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
from sqlalchemy.pool import StaticPool

from database import Base, get_db
from saas_app import app


@pytest.fixture
def client():
    test_engine = create_engine(
        "sqlite://",
        connect_args={"check_same_thread": False},
        poolclass=StaticPool,
    )
    Base.metadata.create_all(test_engine)
    test_sessions = sessionmaker(bind=test_engine, autoflush=False, expire_on_commit=False)
    app.state.test_sessions = test_sessions

    def override_get_db():
        db = test_sessions()
        try:
            yield db
        finally:
            db.close()

    app.dependency_overrides[get_db] = override_get_db
    with TestClient(app) as test_client:
        yield test_client
    app.dependency_overrides.clear()
    del app.state.test_sessions
    Base.metadata.drop_all(test_engine)
    test_engine.dispose()


@pytest.fixture
def signed_in(client):
    csrf = client.get("/auth/csrf").json()["csrf_token"]
    response = client.post(
        "/auth/signup",
        headers={"X-CSRF-Token": csrf},
        json={"tenant_name": "Northwind Traders", "email": "owner@example.com", "password": "a-secure-passphrase-42"},
    )
    assert response.status_code == 200
    return client, client.get("/auth/csrf").json()["csrf_token"]
