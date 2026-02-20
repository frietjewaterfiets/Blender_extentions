bl_info = {
    "name": "IMGFromClipboard (Windows)",
    "author": "ChatGPT",
    "version": (2, 0, 0),
    "blender": (3, 0, 0),
    "location": "3D Viewport > Add (Shift+A) > IMGFromClipboard",
    "description": "Adds an Add-menu entry that imports an image directly from the Windows clipboard",
    "category": "Import-Export",
}

import bpy
import os
import struct
import tempfile
import time
import ctypes
from ctypes import wintypes


# -----------------------------------------------------------------------------
# Persistent temp folder (user can delete manually)
# -----------------------------------------------------------------------------

FOLDER_NAME = "Blender_IMGFromClipboard"
STORAGE_DIR = os.path.join(tempfile.gettempdir(), FOLDER_NAME)


def ensure_storage_dir() -> str:
    os.makedirs(STORAGE_DIR, exist_ok=True)
    return STORAGE_DIR


def make_unique_bmp_path() -> str:
    ensure_storage_dir()
    stamp = time.strftime("%Y%m%d_%H%M%S")
    ms = int((time.time() - int(time.time())) * 1000)
    return os.path.join(STORAGE_DIR, f"clipboard_{stamp}_{ms:03d}.bmp")


# -----------------------------------------------------------------------------
# Windows Clipboard (DIB) helpers
# -----------------------------------------------------------------------------

CF_DIB = 8
CF_DIBV5 = 17

if os.name == "nt":
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    user32.OpenClipboard.argtypes = [wintypes.HWND]
    user32.OpenClipboard.restype = wintypes.BOOL
    user32.CloseClipboard.argtypes = []
    user32.CloseClipboard.restype = wintypes.BOOL
    user32.IsClipboardFormatAvailable.argtypes = [wintypes.UINT]
    user32.IsClipboardFormatAvailable.restype = wintypes.BOOL
    user32.GetClipboardData.argtypes = [wintypes.UINT]
    user32.GetClipboardData.restype = wintypes.HANDLE

    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = wintypes.LPVOID
    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalUnlock.restype = wintypes.BOOL
    kernel32.GlobalSize.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalSize.restype = ctypes.c_size_t


def get_clipboard_dib_bytes():
    """Return DIB bytes from Windows clipboard or None."""
    if os.name != "nt":
        return None

    fmt = None
    if user32.IsClipboardFormatAvailable(CF_DIBV5):
        fmt = CF_DIBV5
    elif user32.IsClipboardFormatAvailable(CF_DIB):
        fmt = CF_DIB
    else:
        return None

    if not user32.OpenClipboard(None):
        return None

    try:
        h = user32.GetClipboardData(fmt)
        if not h:
            return None

        size = kernel32.GlobalSize(h)
        ptr = kernel32.GlobalLock(h)
        if not ptr:
            return None

        try:
            return ctypes.string_at(ptr, size)
        finally:
            kernel32.GlobalUnlock(h)
    finally:
        user32.CloseClipboard()


def dib_to_bmp_file_bytes(dib: bytes) -> bytes:
    """
    Wrap DIB (CF_DIB/CF_DIBV5) into a BMP file by adding BITMAPFILEHEADER.
    """
    if not dib or len(dib) < 40:
        raise ValueError("Clipboard DIB data is invalid or too small.")

    biSize = struct.unpack_from("<I", dib, 0)[0]
    if biSize < 40 or len(dib) < biSize:
        raise ValueError("Invalid DIB header.")

    biBitCount = struct.unpack_from("<H", dib, 14)[0]
    biCompression = struct.unpack_from("<I", dib, 16)[0]
    biClrUsed = struct.unpack_from("<I", dib, 32)[0]

    palette_entries = 0
    if biBitCount <= 8:
        palette_entries = biClrUsed if biClrUsed != 0 else (1 << biBitCount)
    palette_size = palette_entries * 4  # RGBQUAD

    masks_size = 0
    if biCompression in (3, 6):  # BI_BITFIELDS / BI_ALPHABITFIELDS
        if biSize == 40:
            masks_size = 12  # common case

    off_bits = 14 + biSize + palette_size + masks_size
    file_size = 14 + len(dib)

    bfh = struct.pack("<2sIHHI", b"BM", file_size, 0, 0, off_bits)
    return bfh + dib


def save_clipboard_image_to_disk() -> str:
    """
    Saves clipboard image to STORAGE_DIR as BMP and returns the file path.
    Returns empty string if no image.
    """
    dib = get_clipboard_dib_bytes()
    if not dib:
        return ""

    bmp_bytes = dib_to_bmp_file_bytes(dib)
    path = make_unique_bmp_path()
    with open(path, "wb") as f:
        f.write(bmp_bytes)
    return path


def load_image_from_path(path: str):
    try:
        return bpy.data.images.load(path, check_existing=False)
    except Exception:
        return None


# -----------------------------------------------------------------------------
# Placement helpers
# -----------------------------------------------------------------------------

def ensure_object_mode():
    if bpy.context.mode != "OBJECT":
        try:
            bpy.ops.object.mode_set(mode="OBJECT")
        except Exception:
            pass


def add_image_reference(img: bpy.types.Image):
    ensure_object_mode()
    bpy.ops.object.empty_add(type="IMAGE", align="VIEW")
    obj = bpy.context.active_object
    if not obj or obj.type != "EMPTY":
        return None

    obj.empty_display_type = "IMAGE"
    obj.data = img

    if hasattr(obj, "show_in_front"):
        obj.show_in_front = True
    if hasattr(obj, "empty_image_depth"):
        obj.empty_image_depth = "FRONT"

    return obj


def make_material_with_image(img: bpy.types.Image) -> bpy.types.Material:
    mat = bpy.data.materials.new(name=f"Mat_{img.name}")
    mat.use_nodes = True

    nt = mat.node_tree
    nodes = nt.nodes
    links = nt.links

    for n in list(nodes):
        nodes.remove(n)

    tex = nodes.new("ShaderNodeTexImage")
    tex.image = img
    tex.location = (-300, 0)

    bsdf = nodes.new("ShaderNodeBsdfPrincipled")
    bsdf.location = (0, 0)

    out = nodes.new("ShaderNodeOutputMaterial")
    out.location = (250, 0)

    links.new(tex.outputs.get("Color"), bsdf.inputs.get("Base Color"))
    if tex.outputs.get("Alpha") and bsdf.inputs.get("Alpha"):
        links.new(tex.outputs["Alpha"], bsdf.inputs["Alpha"])

    links.new(bsdf.outputs["BSDF"], out.inputs["Surface"])

    # Guard for Blender-version differences
    if hasattr(mat, "blend_method"):
        mat.blend_method = "BLEND"
    if hasattr(mat, "shadow_method"):
        mat.shadow_method = "NONE"

    return mat


def add_mesh_plane_with_image(img: bpy.types.Image):
    ensure_object_mode()
    bpy.ops.mesh.primitive_plane_add(align="VIEW")
    obj = bpy.context.active_object
    if not obj or obj.type != "MESH":
        return None

    # Scale to match aspect ratio (nice default)
    w = max(1, img.size[0])
    h = max(1, img.size[1])
    obj.scale.x *= (w / h)

    mat = make_material_with_image(img)
    if obj.data.materials:
        obj.data.materials[0] = mat
    else:
        obj.data.materials.append(mat)

    return obj


# -----------------------------------------------------------------------------
# Operators & Menus
# -----------------------------------------------------------------------------

TOOLTIP_PATH = STORAGE_DIR  # this will exist after first use; still fine to show

class WM_OT_img_from_clipboard(bpy.types.Operator):
    bl_idname = "wm.img_from_clipboard"
    bl_label = "IMGFromClipboard"
    bl_description = (
        "Imports an image from the Windows clipboard.\n"
        f"Images are saved to: {TOOLTIP_PATH}"
    )
    bl_options = {"REGISTER", "UNDO"}

    mode: bpy.props.EnumProperty(
        name="Import As",
        items=[
            ("REFERENCE", "Reference", "Create an Image Empty (Reference)"),
            ("MESH", "Mesh", "Create a Plane with the image as material"),
        ],
        default="REFERENCE",
    )

    def execute(self, context):
        if os.name != "nt":
            self.report({"WARNING"}, "IMGFromClipboard is Windows-only.")
            return {"CANCELLED"}

        saved_path = save_clipboard_image_to_disk()
        if not saved_path:
            self.report({"WARNING"}, "No image found in clipboard.")
            return {"CANCELLED"}

        img = load_image_from_path(saved_path)
        if not img:
            self.report({"WARNING"}, f"Failed to load saved image: {saved_path}")
            return {"CANCELLED"}

        if self.mode == "MESH":
            obj = add_mesh_plane_with_image(img)
        else:
            obj = add_image_reference(img)

        if not obj:
            self.report({"WARNING"}, "Failed to create object.")
            return {"CANCELLED"}

        self.report({"INFO"}, f"Imported from clipboard ({self.mode}). Saved at: {saved_path}")
        return {"FINISHED"}


class VIEW3D_MT_img_from_clipboard(bpy.types.Menu):
    bl_label = "IMGFromClipboard"
    bl_description = f"Images are saved to: {TOOLTIP_PATH}"

    def draw(self, context):
        layout = self.layout
        op = layout.operator("wm.img_from_clipboard", text="Reference", icon="IMAGE_REFERENCE")
        op.mode = "REFERENCE"

        op = layout.operator("wm.img_from_clipboard", text="Mesh", icon="MESH_PLANE")
        op.mode = "MESH"


def draw_img_from_clipboard_in_add_menu(self, context):
    # Adds: Add > IMGFromClipboard > (Reference / Mesh)
    self.layout.menu("VIEW3D_MT_img_from_clipboard", icon="IMAGE_DATA")


classes = (
    WM_OT_img_from_clipboard,
    VIEW3D_MT_img_from_clipboard,
)


def register():
    ensure_storage_dir()
    for c in classes:
        bpy.utils.register_class(c)
    bpy.types.VIEW3D_MT_add.append(draw_img_from_clipboard_in_add_menu)


def unregister():
    bpy.types.VIEW3D_MT_add.remove(draw_img_from_clipboard_in_add_menu)
    for c in reversed(classes):
        bpy.utils.unregister_class(c)