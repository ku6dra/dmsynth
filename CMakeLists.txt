cmake_minimum_required(VERSION 3.15)
project(DMSynth LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

add_executable(dm_synth
    src/main.cpp
    src/midi_input.cpp
    src/dm_synth.cpp
    src/cli_args.cpp
)

target_link_libraries(dm_synth PRIVATE
    winmm
    dsound
    dxguid
    ole32
    uuid
    avrt
)


if(MSVC)
    target_compile_options(dm_synth PRIVATE /utf-8)
    target_compile_options(dm_synth PRIVATE
        $<$<CONFIG:Release>:/O2 /Ob3 /GL /GS- /fp:fast>
    )
    target_link_options(dm_synth PRIVATE
        $<$<CONFIG:Release>:/LTCG /OPT:REF /OPT:ICF>
    )
endif()

target_compile_definitions(dm_synth PRIVATE
    _WIN32_WINNT=0x0A00
    UNICODE
    _UNICODE
)
