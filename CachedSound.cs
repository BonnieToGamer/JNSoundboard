﻿using System.Collections.Generic;
using System.Linq;
using NAudio.Wave;

// https://mark-dot-net.blogspot.de/2014/02/fire-and-forget-audio-playback-with.html
// Mark Heath 2014

namespace JNSoundboard
{
    class CachedSound
    {
        public float[] AudioData { get; private set; }
        public WaveFormat WaveFormat { get; private set; }

        public CachedSound(string audioFileName)
        {
            using (AudioFileReader audioFileReader = new AudioFileReader(audioFileName))
            {
                // TODO: could add resampling in here if required
                WaveFormat = audioFileReader.WaveFormat;
                List<float> wholeFile = new List<float>((int)(audioFileReader.Length / 4));
                float[] readBuffer = new float[audioFileReader.WaveFormat.SampleRate * audioFileReader.WaveFormat.Channels];
                int samplesRead;

                while ((samplesRead = audioFileReader.Read(readBuffer, 0, readBuffer.Length)) > 0)
                {
                    wholeFile.AddRange(readBuffer.Take(samplesRead));
                }

                AudioData = wholeFile.ToArray();
            }
        }
    }
}
