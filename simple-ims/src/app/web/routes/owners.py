from fastapi import APIRouter, Request, Depends, Form, status
from fastapi.responses import RedirectResponse
from app.web.templating import create_templates
from sqlalchemy.orm import Session
from app.db import get_db
from app.services.owners import OwnerService
from app.schemas.owner import OwnerCreate, OwnerUpdate

router = APIRouter(prefix="/owners")
templates = create_templates()


@router.get("/")
def list_owners(request: Request, error: str | None = None, db: Session = Depends(get_db)):
    owners = OwnerService.list_owners(db)
    return templates.TemplateResponse(
        request, "owners/list.html", {"owners": owners, "error": error}
    )


@router.get("/new")
def new_owner(request: Request):
    return templates.TemplateResponse(request, "owners/form.html", {"owner": None})


@router.post("/")
def create_owner(
    request: Request,
    name: str = Form(),
    code: str = Form(),
    enabled: bool = Form(default=True),
    db: Session = Depends(get_db),
):
    data = OwnerCreate(name=name, code=code, enabled=enabled)
    OwnerService.create_owner(db, data)
    return RedirectResponse(url="/owners/", status_code=status.HTTP_303_SEE_OTHER)


@router.get("/{owner_id}/edit")
def edit_owner(request: Request, owner_id: int, db: Session = Depends(get_db)):
    owner = OwnerService.get_owner(db, owner_id)
    if not owner:
        return RedirectResponse(url="/owners/", status_code=status.HTTP_302_FOUND)
    return templates.TemplateResponse(
        request, "owners/form.html", {"owner": owner}
    )


@router.post("/{owner_id}/edit")
def update_owner(
    request: Request,
    owner_id: int,
    name: str = Form(),
    code: str = Form(),
    enabled: bool = Form(default=True),
    db: Session = Depends(get_db),
):
    owner = OwnerService.get_owner(db, owner_id)
    if not owner:
        return RedirectResponse(url="/owners/", status_code=status.HTTP_302_FOUND)
    data = OwnerUpdate(name=name, code=code, enabled=enabled)
    OwnerService.update_owner(db, owner, data)
    return RedirectResponse(url="/owners/", status_code=status.HTTP_303_SEE_OTHER)


@router.post("/{owner_id}/delete")
def delete_owner(
    request: Request,
    owner_id: int,
    db: Session = Depends(get_db),
):
    owner = OwnerService.get_owner(db, owner_id)
    if not owner:
        return RedirectResponse(url="/owners/", status_code=status.HTTP_302_FOUND)
    try:
        OwnerService.delete_owner(db, owner)
    except ValueError as exc:
        owners = OwnerService.list_owners(db)
        return templates.TemplateResponse(
            request,
            "owners/list.html",
            {"owners": owners, "error": str(exc)},
            status_code=status.HTTP_400_BAD_REQUEST,
        )
    return RedirectResponse(url="/owners/", status_code=status.HTTP_303_SEE_OTHER)
