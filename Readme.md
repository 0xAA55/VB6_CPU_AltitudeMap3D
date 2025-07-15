# AltitudeMap

A real-time CPU-based ray-traced terrain map renderer written in Visual Basic 6.0

## 语言 Language

[简体中文](Readme-CN.md) | English

## Introduction
Raymarch ray-tracing technology is a method that, for each pixel, calculates the line-of-sight direction and then performs a limited number of iterative steps to find the intersection point between the line of sight and the height map represented by the displacement texture. It is used in computer graphics as a simplified ray-tracing solution.

Due to the fact that the complexity of this algorithm can be optimized to a very low level, even the ancient programming language Visual Basic 6.0 can achieve smooth real-time ray-tracing rendering relying solely on the CPU.

This ray-tracing algorithm is only used to calculate the intersection points between the line of sight and objects, and does not provide physical optical calculations (such as diffuse reflection statistics). It is a new rendering method that replaces traditional triangular face geometry.

## Example Effect
![Example](record.gif)

## Features
- Uses a thread pool for rendering.
- Does not use any SIMD acceleration instruction sets, relying purely on VB6 native single-precision floating-point calculations.
- Keyboard and mouse interaction: W/A/S/D for movement, spacebar for jumping.
- The rendering method is Raymarch ray-tracing, but the shading model uses the traditional N·L method.
- Infinitely repeated map size with no boundaries.
- Supports configuration of rendering parameters via a configuration file.
- Press the P key once to automatically start recording the screen to the record.avi file, and press P again to stop recording.
  - Note: The recording duration is limited, and the resulting AVI file is uncompressed, resulting in a large file size. Use this function with caution.

## Raymarch Ray-Tracing
The Raymarch ray-tracing algorithm uses various methods to calculate the step size of the line-of-sight vector. For specific geometries, a Signed Distance Function (SDF) designed for the geometry can be used to calculate distance values. It traverses all geometries in the entire scene and selects the minimum step size for stepping.

Our current code example uses a different algorithm: intersection calculation for displacement maps. Each pixel of the displacement map represents the height value of that point relative to the plane.

How to design a reasonable step size for displacement? The most straightforward algorithm is to fix a step size and then determine whether the line of sight is embedded in the object (terrain).

But there are also smarter algorithms, as shown in the diagram:
![Terrain cone stepping algorithm](demo.PNG)

We traverse each pixel of the displacement map and then check whether all surrounding pixels are higher than the current pixel. If so, we calculate the steepness based on the distance of that pixel and finally store the maximum steepness to create a K-map, where each pixel value represents the taper of a cone emitted vertically upward from that pixel.

Then, when the line of sight steps forward, it samples the K-map from the 2D direction of the current line of sight to obtain the cone's taper, and then intersects the line of sight with the cone to achieve an "elegant step".

