#pragma once

#include <string>

constexpr int SPECTROGRAM_WIDTH = 1280;
constexpr int SPECTROGRAM_HEIGHT = 720;
constexpr const wchar_t* SPECTROGRAM_COLOR = L"intensity";
constexpr const wchar_t* SPECTROGRAM_MODE = L"combined";
constexpr const wchar_t* SPECTROGRAM_SCALE = L"log";
constexpr const wchar_t* SPECTROGRAM_UNSHARP = L"5:5:0.8:5:5:0.0";
constexpr unsigned long FFMPEG_TIMEOUT_MS = 30000;

bool GenerateAndShowSpectrogram(const std::wstring& audioFile);
