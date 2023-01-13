# Audio Extractor
from pathlib import Path
import glob, os, shutil
from pydub import AudioSegment
from tqdm import tqdm
import zipfile 

def welcome_message():
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi welcome to Audio_Extractor.py!')
    print(f'This is a quick program taking and compiling audio files from a .pptx format in the same directory')
    if input("Enter y to continue, anything else to exit\n") == 'y':

        # copy all .pptx into an ae_copies folder
        if os.path.exists("ae_copies"):
            shutil.rmtree("ae_copies")
        os.mkdir("ae_copies")

        for file in glob.glob("*.pptx"):
            print("powerpoint found {f}".format(f=file))
            shutil.copyfile('{soc}'.format(soc=file), '{dest}'.format(dest='ae_copies/' + file))

        # perform the extractions on all files
        for file in glob.glob('ae_copies/*.pptx'):
            extract_audio(file)

    else:
        exit()


def extract_audio(file):
    # Makes a copy and extracts audio via zip file
    # output a list of all audio AND the OG file name to extract_files
    p = Path(file)
    zip_p = p.rename(p.with_suffix('.zip'))

    zip_p = Path.joinpath(zip_p,"ppt")
    print(zip_p)
    print(os.path.exists("{path_name}".format(path_name=zip_p)))
    for file in glob.glob("{path_name}".format(path_name=zip_p)):
        print("Found file of name {n}".format(n=file))
    # Go in and send the list of all mp4 files


def compile_audio(audio_clip_paths, output_path):
    # Takes the output from extract_files and compiles into one
    # Credit goes to Abdou Rockikz @ ThePythonCode.Com
    clips = []

    audio_clip_paths = tqdm(audio_clip_paths, "Reading Audio File")
    for path in audio_clip_paths:
        extension = get_file_type(path)
        clip = AudioSegment.from_file(path, extension)
        clips.append(clip)

    final_clip = clips[0]
    range_loop = tqdm(list(range(1, len(clips))), "Concatenating Audio")
    for i in range_loop:
        final_clip = final_clip + clips[i]

    final_clip_extension = get_file_type(output_path)
    print("Exporting consolidated audio file to {o}".format(o=output_path))
    final_clip.export(output_path, format=final_clip_extension)


def get_file_type(filename):
    return os.path.splitext(filename)[1].lstrip(".")


if __name__ == '__main__':
    welcome_message()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
