import os
import cv2
import argparse
import torch
import numpy as np

from utils import IMG_SIZE, bilinear_unwarping, load_model

ckpt_path='./UVDoc/model'

def unwarp_and_save(ckpt_path, src_path, dst_path, img_size):
    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    model = load_model(ckpt_path)
    model.to(device)
    model.eval()

    img = cv2.imread(src_path)
    if img is None:
        print("⚠️ 无法读取图片，跳过:", src_path)
        return False

    img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB).astype(np.float32) / 255.0
    inp = torch.from_numpy(cv2.resize(img_rgb, img_size).transpose(2, 0, 1)).unsqueeze(0).to(device)

    with torch.no_grad():
        point_positions2D, _ = model(inp)

    h, w = img.shape[:2]
    unwarped = bilinear_unwarping(
        warped_img=torch.from_numpy(img_rgb.transpose(2,0,1)).unsqueeze(0).to(device),
        point_positions=torch.unsqueeze(point_positions2D[0], dim=0),
        img_size=(w, h)
    )
    unwarped = (unwarped[0].cpu().numpy().transpose(1,2,0) * 255).astype(np.uint8)
    unwarped_bgr = cv2.cvtColor(unwarped, cv2.COLOR_RGB2BGR)

    os.makedirs(os.path.dirname(dst_path), exist_ok=True)
    cv2.imwrite(dst_path, unwarped_bgr)
    print("✅ saved:", dst_path)
    return True

def batch_process(root_in, root_out, ckpt_path, exts={".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"}):
    for dirpath, dirnames, filenames in os.walk(root_in):
        rel = os.path.relpath(dirpath, root_in)
        for fn in filenames:
            ext = os.path.splitext(fn)[1].lower()
            if ext in exts:
                src = os.path.join(dirpath, fn)
                dst = os.path.join(root_out, rel, os.path.splitext(fn)[0] + "_unwarp.png")
                unwarp_and_save(ckpt_path, src, dst, IMG_SIZE)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Batch unwarp images")
    parser.add_argument("--ckpt", type=str, required=True, help="path of model checkpoint")
    parser.add_argument("--src", type=str, default="./img", help="root folder for input images")
    parser.add_argument("--dst", type=str, default="./img_edit", help="root folder for output images")
    args = parser.parse_args()

    batch_process(args.src, args.dst, args.ckpt)
