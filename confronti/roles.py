ILLUMIA_OPERATOR_GROUP = "Operatori Illumia"


def is_illumia_operator(user):
    if not getattr(user, "is_authenticated", False):
        return False
    return user.groups.filter(name=ILLUMIA_OPERATOR_GROUP).exists()


def is_internal_user(user):
    return bool(
        getattr(user, "is_authenticated", False)
        and (getattr(user, "is_staff", False) or getattr(user, "is_superuser", False) or is_illumia_operator(user))
    )
