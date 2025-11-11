import numpy as np
import torch
from torchvision import models, transforms
import matplotlib.pyplot as plt
import cv2
from scipy.ndimage import gaussian_filter

# بارگذاری تصویر
image_path = '*.jpg'  # مسیر عکست رو جایگزین کن
image = cv2.imread(image_path)
image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)

# ذخیره اندازه اصلی تصویر
original_height, original_width = image.shape[:2]

# نمایش تصویر اصلی
plt.imshow(image)
plt.title('تصویر اصلی')
plt.axis('off')
plt.show()

# بارگذاری مدل DeepLabV3
model = models.segmentation.deeplabv3_resnet101(weights=models.segmentation.DeepLabV3_ResNet101_Weights.DEFAULT)
model.eval()

# آماده‌سازی تصویر برای مدل
preprocess = transforms.Compose([
    transforms.ToPILImage(),
    transforms.Resize((520, 520)),
    transforms.ToTensor(),
    transforms.Normalize(mean=[0.485, 0.456, 0.406], std=[0.229, 0.224, 0.225]),
])

input_tensor = preprocess(image)
input_batch = input_tensor.unsqueeze(0)

if torch.cuda.is_available():
    input_batch = input_batch.to('cuda')
    model.to('cuda')

with torch.no_grad():
    output = model(input_batch)['out'][0]
output_predictions = output.argmax(0).cpu().numpy()

# ماسک شخص (کلاس 15 برای person)
person_mask = (output_predictions == 15).astype(np.uint8)

# resize ماسک به اندازه اصلی تصویر
person_mask = cv2.resize(person_mask, (original_width, original_height), interpolation=cv2.INTER_NEAREST).astype(np.bool_)

# نمایش ماسک شخص
plt.imshow(person_mask, cmap='gray')
plt.title('ماسک شخص')
plt.axis('off')
plt.show()

# تابع تبدیل RGB به HSV
def rgb_to_hsv(rgb):
    rgb = rgb.astype(np.float32) / 255.0
    cmax = np.max(rgb, axis=-1)
    cmin = np.min(rgb, axis=-1)
    delta = cmax - cmin
    
    h = np.zeros_like(cmax)
    s = np.zeros_like(cmax)
    v = cmax
    
    mask = delta > 0
    r, g, b = rgb[..., 0], rgb[..., 1], rgb[..., 2]
    
    mask_r = (cmax == r) & mask
    h[mask_r] = (60 * ((g[mask_r] - b[mask_r]) / delta[mask_r]) % 360) / 360
    
    mask_g = (cmax == g) & mask
    h[mask_g] = (60 * ((b[mask_g] - r[mask_g]) / delta[mask_g] + 120) % 360) / 360
    
    mask_b = (cmax == b) & mask
    h[mask_b] = (60 * ((r[mask_b] - g[mask_b]) / delta[mask_b] + 240) % 360) / 360
    
    s[mask] = delta[mask] / cmax[mask]
    return np.stack([h, s, v], axis=-1)

# تبدیل به HSV و ماسک پوست
hsv = rgb_to_hsv(image)
h, s, v = hsv[..., 0], hsv[..., 1], hsv[..., 2]
skin_mask = (h >= 0.0) & (h <= 0.12) & (s >= 0.15) & (s <= 0.7) & (v >= 0.45) & (v <= 1.0)
skin_mask = skin_mask & person_mask

# نمایش ماسک پوست
plt.imshow(skin_mask, cmap='gray')
plt.title('ماسک پوست')
plt.axis('off')
plt.show()

# ماسک لباس
cloth_mask = person_mask & ~skin_mask

# محاسبه رنگ متوسط پوست
skin_pixels = image[skin_mask]
if skin_pixels.size > 0:
    avg_skin_color = np.mean(skin_pixels, axis=0).astype(np.uint8)
else:
    avg_skin_color = np.array([210, 170, 140], dtype=np.uint8)

# ایجاد تصویر اولیه بدون لباس
nude_image = image.copy()
nude_image[cloth_mask] = avg_skin_color

# بهبود با Gaussian blur
blurred_skin = np.zeros_like(image, dtype=np.float32)
for c in range(3):
    blurred_skin[..., c] = gaussian_filter(image[..., c], sigma=2)
nude_image[cloth_mask] = blurred_skin[cloth_mask].astype(np.uint8)

# استفاده از Inpainting برای طبیعی‌تر کردن
cloth_mask_cv = cloth_mask.astype(np.uint8) * 255
nude_image = cv2.inpaint(nude_image, cloth_mask_cv, inpaintRadius=7, flags=cv2.INPAINT_TELEA)

# نمایش تصویر نهایی
plt.imshow(nude_image)
plt.title('تصویر بدون لباس (بهبودیافته)')
plt.axis('off')
plt.show()

# ذخیره تصویر
cv2.imwrite('nude_image_improved.jpg', cv2.cvtColor(nude_image, cv2.COLOR_RGB2BGR))