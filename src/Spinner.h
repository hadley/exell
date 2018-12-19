#ifndef READXL_SPINNER_
#define READXL_SPINNER_

#include <RProgress.h>

class Spinner {
  bool progress_;
  RProgress::RProgress pb_;

public:
  Spinner(bool progress = true):
  progress_(progress)
  {
    if (progress_) {
      pb_ = RProgress::RProgress(":spin");
      pb_.set_total(1);
      pb_.set_show_after(2);
    }
  }
  void spin()   { if (progress_) pb_.update(0.5); }
  void finish() { if (progress_) pb_.update(1);   }
  ~Spinner() { if (this->progress_) this->finish(); }
};

#endif
