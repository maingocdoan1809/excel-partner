import type { ResolvedInsertImageInput } from "../../../protocol/src/index.js";
import { ToolFailure } from "./common.js";

export interface Size { width: number; height: number }
export interface Box extends Size { left: number; top: number }

/** Office draws pictures at 96 pixels to the inch and measures them in points. */
export const POINTS_PER_PIXEL = 0.75;

const round = (value: number): number => Math.round(value * 100) / 100;

/** The picture's own size in points, from the pixel size the bridge read. Null when the file did not say. */
export function naturalSize(input: Pick<ResolvedInsertImageInput, "naturalWidth" | "naturalHeight">): Size | null {
  const { naturalWidth, naturalHeight } = input;
  return naturalWidth && naturalHeight ? { width: naturalWidth * POINTS_PER_PIXEL, height: naturalHeight * POINTS_PER_PIXEL } : null;
}

/**
 * The size a picture gets, in points, never stretched: with both `width` and `height` it is fitted inside that box; with one the
 * other follows; with neither it keeps `natural` size, shrunk to fit `limit` when it would be larger (never enlarged).
 */
export function fitImage(natural: Size, request: { width?: number; height?: number }, limit?: Size | null): Size {
  if (!(natural.width > 0) || !(natural.height > 0)) throw new ToolFailure("office_error", "the picture has no size, so it cannot be fitted");
  let scale: number;
  if (request.width !== undefined && request.height !== undefined) scale = Math.min(request.width / natural.width, request.height / natural.height);
  else if (request.width !== undefined) scale = request.width / natural.width;
  else if (request.height !== undefined) scale = request.height / natural.height;
  else scale = limit ? Math.min(1, limit.width / natural.width, limit.height / natural.height) : 1;
  return { width: round(natural.width * scale), height: round(natural.height * scale) };
}

/** `size` centred inside `area`. */
export function centreIn(area: Box, size: Size): Box {
  return { ...size, left: round(area.left + (area.width - size.width) / 2), top: round(area.top + (area.height - size.height) / 2) };
}
