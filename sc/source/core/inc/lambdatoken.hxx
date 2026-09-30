/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
/*
 * This file is part of the LibreOffice project.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 */

#pragma once

#include <formula/token.hxx>
#include <rtl/ustring.hxx>
#include <tokenarray.hxx>

#include <memory>
#include <utility>
#include <vector>

/** The value of a LAMBDA(...) expression: parameter names and a body.

    The body is a range of RPN code in a copy of the token array the LAMBDA
    was evaluated in, so names bound by an enclosing LET or LAMBDA have
    already been replaced by their values. Helper functions such as MAP bind
    the parameters and evaluate the body once per element.

    The token has no stack type of its own (svUnknown): it is only meaningful
    as an argument to a function that expects a LAMBDA. Anywhere else it is an
    error, as in Excel where an uncalled LAMBDA gives #CALC!. */
class ScLambdaToken final : public formula::FormulaToken
{
    std::shared_ptr<const ScTokenArray> mpCode;
    short mnBodyStart;
    short mnBodyEnd;
    std::vector<OUString> maNames;

public:
    ScLambdaToken(std::shared_ptr<const ScTokenArray> pCode, short nBodyStart, short nBodyEnd,
                  std::vector<OUString> aNames)
        : FormulaToken(formula::svUnknown, ocLambda)
        , mpCode(std::move(pCode))
        , mnBodyStart(nBodyStart)
        , mnBodyEnd(nBodyEnd)
        , maNames(std::move(aNames))
    {
    }

    const ScTokenArray& GetCode() const { return *mpCode; }
    short GetBodyStart() const { return mnBodyStart; }
    short GetBodyEnd() const { return mnBodyEnd; }
    const std::vector<OUString>& GetNames() const { return maNames; }

    virtual FormulaToken* Clone() const override { return new ScLambdaToken(*this); }
    virtual bool operator==(const FormulaToken& rToken) const override
    {
        const ScLambdaToken* pOther = dynamic_cast<const ScLambdaToken*>(&rToken);
        return pOther && pOther->mpCode == mpCode && pOther->mnBodyStart == mnBodyStart
               && pOther->mnBodyEnd == mnBodyEnd && pOther->maNames == maNames;
    }
};

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
