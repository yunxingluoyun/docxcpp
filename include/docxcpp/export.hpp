#pragma once

#if defined(_WIN32) || defined(__CYGWIN__)
#  if defined(DOCXCPP_STATIC_DEFINE)
#    define DOCXCPP_API
#  elif defined(DOCXCPP_BUILDING_LIBRARY)
#    define DOCXCPP_API __declspec(dllexport)
#  else
#    define DOCXCPP_API __declspec(dllimport)
#  endif
#elif defined(__GNUC__) && __GNUC__ >= 4
#  if defined(DOCXCPP_STATIC_DEFINE)
#    define DOCXCPP_API
#  else
#    define DOCXCPP_API __attribute__((visibility("default")))
#  endif
#else
#  define DOCXCPP_API
#endif
