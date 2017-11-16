#ifndef AUDITXLS_GLOBAL_H
#define AUDITXLS_GLOBAL_H

#ifndef QT_IDE
#define VS_IDE
#endif

#ifdef QT_IDE
#include <QString>

#define OE_DECLARE_PRIVATE(Class)  Q_DECLARE_PRIVATE(Class)

#define OE_DECLARE_PRIVATE_D(Dptr, Class) Q_DECLARE_PRIVATE_D(Dptr, Class)

#define OE_DECLARE_PUBLIC(Class)  Q_DECLARE_PUBLIC(Class)

#define OE_D(Class) Q_D(Class)
#define OE_Q(Class) Q_Q(Class)

#define OE_DISABLE_COPY(Class) Q_DISABLE_COPY(Class)

#define OE_NULLPTR Q_NULLPTR

#elif defined(VS_IDE)
#pragma comment(lib, "libxl.lib")

#define _CRT_SECURE_NO_WARNINGS


#if __cplusplus>201111L
# define Q_DECL_EQ_DELETE = delete
#else
# define Q_DECL_EQ_DELETE
#endif

template <typename T> static inline T *qGetPtrHelper(T *ptr) { return ptr; }
template <typename Wrapper> static inline typename Wrapper::pointer qGetPtrHelper(const Wrapper &p) { return p.data(); }

#define OE_DECLARE_PRIVATE(Class) \
    inline Class##Private* d_func() { return reinterpret_cast<Class##Private *>(qGetPtrHelper(d_ptr)); } \
    inline const Class##Private* d_func() const { return reinterpret_cast<const Class##Private *>(qGetPtrHelper(d_ptr)); } \
    friend class Class##Private;

#define OE_DECLARE_PRIVATE_D(Dptr, Class) \
    inline Class##Private* d_func() { return reinterpret_cast<Class##Private *>(Dptr); } \
    inline const Class##Private* d_func() const { return reinterpret_cast<const Class##Private *>(Dptr); } \
    friend class Class##Private;

#define OE_DECLARE_PUBLIC(Class)                                    \
    inline Class* q_func() { return static_cast<Class *>(q_ptr); } \
    inline const Class* q_func() const { return static_cast<const Class *>(q_ptr); } \
    friend class Class;

#define OE_D(Class) Class##Private * const d = d_func()
#define OE_Q(Class) Class * const q = q_func()

#define OE_DISABLE_COPY(Class) \
    Class(const Class &) Q_DECL_EQ_DELETE;\
    Class &operator=(const Class &) Q_DECL_EQ_DELETE;

#define OE_NULLPTR nullptr
#endif

#ifndef __xls_in
#define __xls_in
#endif

#ifndef __xls_out
#define __xls_out
#endif


#if defined(AUDITXLS_LIBRARY)
#  define AUDITXLSSHARED_EXPORT __declspec(dllexport)
#else
#  define AUDITXLSSHARED_EXPORT __declspec(dllimport)
#endif

#endif // AUDITXLS_GLOBAL_H
