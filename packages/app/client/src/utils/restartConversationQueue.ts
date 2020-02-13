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
import { Activity } from 'botframework-schema';

export class ConversationQueue {
  private activities: Activity[] = [];
  private conversationId: string = '';
  private restartFlow: Map<string, Map<string, number>>;
  private replayActivities: Map<string, Activity>;
  private activitiesToBeProcessed: string[];
  private createObjectUrlFromWindow: Function;

  constructor(conversationId: string, originalActivities: Activity[], createObjectUrl: Function) {
    this.activities = originalActivities;
    this.conversationId = conversationId;
    this.buildActivityQueue = this.buildActivityQueue.bind(this);
    this.traverseRestartFlowMap = this.traverseRestartFlowMap.bind(this);
    this.createObjectUrlFromWindow = createObjectUrl;
    this.buildActivityQueue();
  }

  private static dataURLtoFile(dataurl: string, filename: string) {
    var arr = dataurl.split(','),
      mime = arr[0].match(/:(.*?);/)[1],
      bstr = atob(arr[1]),
      n = bstr.length,
      u8arr = new Uint8Array(n);

    while (n--) {
      u8arr[n] = bstr.charCodeAt(n);
    }

    return new File([u8arr], filename, { type: mime });
  }

  public buildActivityQueue(): void {
    this.restartFlow = new Map<string, Map<string, number>>();
    this.replayActivities = new Map<string, Activity>();
    const activityReplyCount = new Map<string, number>();

    this.activities.forEach((activity: Activity) => {
      if (activity.from.role === 'user') {
        const replayActivity: Activity = {
          ...activity,
          channelData: {
            ...activity.channelData,
            originalActivityId: activity.id,
          },
          conversation: {
            ...activity.conversation,
            id: this.conversationId,
          },
        };
        this.replayActivities.set(activity.id, replayActivity);
        this.restartFlow.set(activity.id, new Map(activityReplyCount));
      } else {
        activityReplyCount.set(activity.id, (activityReplyCount.get(activity.id) || 0) + 1);
      }
    });
  }

  public traverseRestartFlowMap(incomingActivityId: string) {
    [...this.restartFlow.keys()].forEach((activityId: string) => {
      const value: Map<string, number> = this.restartFlow.get(activityId);
      let ctOfActivity: number | undefined = value.get(incomingActivityId);
      if (ctOfActivity) {
        ctOfActivity--;
        if (ctOfActivity === 0) {
          value.delete(activityId);
        } else {
          value.set(activityId, ctOfActivity);
        }
      }

      if ([...value.keys()].length === 0) {
        this.activitiesToBeProcessed.push(activityId);
      }
    });
  }

  public getNextActivityInQueue() {
    if (this.activitiesToBeProcessed.length > 0) {
      const activityId: string = this.activitiesToBeProcessed.shift();
      const activity: Activity = this.replayActivities.get(activityId);
      if (activity.attachments && activity.attachments.length >= 1) {
        const mutatedAttachments = activity.attachments.map(attachment => {
          const fileFormat: File = ConversationQueue.dataURLtoFile(attachment.contentUrl, attachment.name);
          return {
            ...attachment,
            contentUrl: this.createObjectUrlFromWindow(fileFormat),
          };
        });
        activity.attachments = mutatedAttachments;
      }
    } else {
      return undefined;
    }
  }
}
