//
// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license.
//
// Microsoft Bot Framework: http://botframework.com
//
// Bot Framework Emulator Github:
// https://github.com/Microsoft/BotFramwork-Emulator
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//

import * as React from 'react';
import { SharedConstants } from '@bfemulator/app-shared';
import { Activity } from 'botframework-schema';

import { areActivitiesEqual } from '../../../../../utils';

import { ActivityWrapper } from './activityWrapper';

export interface OuterActivityWrapperProps {
  card?: any;
  children?: any;
  highlightedActivities?: Activity[];
  documentId: string;
  onContextMenu?: (event: React.MouseEvent<HTMLElement>) => void;
  onItemRendererClick?: (event: React.MouseEvent<HTMLElement>) => void;
  onItemRendererKeyDown?: (event: React.KeyboardEvent<HTMLElement>) => void;
  onRestartConversationFromActivityClick?: (documentId: string, activity: Activity) => void;
}

export class OuterActivityWrapper extends React.Component<OuterActivityWrapperProps, {}> {
  public render() {
    const { card, children, onContextMenu, onItemRendererClick, onItemRendererKeyDown } = this.props;

    return (
      <ActivityWrapper
        activity={card.activity}
        data-activity-id={card.activity.id}
        onClick={onItemRendererClick}
        onKeyDown={onItemRendererKeyDown}
        onContextMenu={onContextMenu}
        isSelected={this.shouldBeSelected(card.activity)}
        isUserActivity={this.isUserActivity(card.activity)}
        onRestartConversationFromActivityClick={this.propsBoundRestartActivityHandler}
      >
        {children}
      </ActivityWrapper>
    );
  }

  private propsBoundRestartActivityHandler = () => {
    this.props.onRestartConversationFromActivityClick(this.props.documentId, this.props.card.activity);
  };

  private isUserActivity(activity: Activity) {
    return (
      activity.from.role === SharedConstants.Activity.FROM_USER_ROLE && !activity.replyToId && activity.channelData
    );
  }

  private shouldBeSelected(subject: Activity): boolean {
    return this.props.highlightedActivities.some(activity => areActivitiesEqual(activity, subject));
  }
}
